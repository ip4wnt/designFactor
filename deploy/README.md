# Развёртывание и страховка от зависания

Прод: VM 4 vCPU / 16 ГБ, Ubuntu, nginx → `127.0.0.1:8811` (uvicorn). В сентябре
2026 на конфигурации 2 vCPU / 2 ГБ без swap три параллельные генерации
с LibreOffice исчерпали память: ядро ушло в вытеснение кэша, sshd и nginx
перестали отвечать, помог только аппаратный reset. Отсюда пять слоёв защиты.

| Слой | Файл | Что делает |
|---|---|---|
| 1. Очередь задач | `backend/.env` → `JOB_CONCURRENCY=3` | одновременно идут не более N тяжёлых стадий (генерация трёх вариантов, пересборка, аудит); остальные ждут в `queued`. Ориентир — ~3 ГБ на задачу. LibreOffice дополнительно сериализован и убивается по таймауту (`app/pipeline/soffice.py`) |
| 2. cgroup-лимит | `designfactor.service` | `MemoryHigh=10G` (мягкое вытеснение только нашего сервиса), `MemoryMax=12G` (OOM-kill внутри cgroup), `Restart=always`, `RestartSec=3`. ОС, sshd, nginx и агенты хостинга (~0,5 ГБ) вне лимита |
| 3. Swap-буфер | `/swapfile` 4 ГБ, `vm.swappiness=10` | пик выше лимита превращается в замедление, а не в трэшинг |
| 4. earlyoom | `earlyoom.default` → `/etc/default/earlyoom` | при свободной памяти < 5 % убивает самый крупный процесс за миллисекунды, предпочитая soffice/uvicorn и не трогая sshd/nginx |
| 5. Самопроверка | `designfactor-healthcheck.{sh,service,timer}` | раз в минуту `curl /docs`; три неответа подряд → `systemctl restart designfactor`; в журнал пишется `free -m` |

## Установка

```bash
cd /home/pankratov/designfactor

# 1. очередь
grep -q '^JOB_CONCURRENCY=' backend/.env || echo 'JOB_CONCURRENCY=3' >> backend/.env

# 2. unit сервиса
sudo cp deploy/designfactor.service /etc/systemd/system/
sudo systemctl daemon-reload && sudo systemctl enable designfactor

# 3. swap
sudo fallocate -l 4G /swapfile && sudo chmod 600 /swapfile
sudo mkswap /swapfile && sudo swapon /swapfile
echo '/swapfile none swap sw 0 0' | sudo tee -a /etc/fstab
echo 'vm.swappiness=10' | sudo tee /etc/sysctl.d/90-swappiness.conf && sudo sysctl --system

# 4. earlyoom
sudo apt-get install -y earlyoom
sudo cp deploy/earlyoom.default /etc/default/earlyoom
sudo systemctl enable --now earlyoom && sudo systemctl restart earlyoom

# 5. самопроверка
sudo cp deploy/designfactor-healthcheck.sh /usr/local/bin/
sudo cp deploy/designfactor-healthcheck.service deploy/designfactor-healthcheck.timer /etc/systemd/system/
sudo systemctl daemon-reload && sudo systemctl enable --now designfactor-healthcheck.timer

sudo systemctl restart designfactor
```

## Проверка

```bash
systemctl show designfactor -p MemoryHigh -p MemoryMax -p Restart
swapon --show && cat /proc/sys/vm/swappiness
systemctl is-active earlyoom designfactor-healthcheck.timer
journalctl -u designfactor-healthcheck --since '10 min ago' --no-pager   # строки ok/FAIL с памятью
journalctl -u earlyoom -b --no-pager | tail                              # что убивал earlyoom
```

Для другого объёма памяти масштабируйте: `MemoryHigh` ≈ 60 %, `MemoryMax` ≈ 75 %
RAM, `JOB_CONCURRENCY` ≈ RAM / 3 ГБ (не меньше 1), swap — 25 % RAM.

nginx: статика `frontend/` из `/var/www/designfactor/`, proxy `/jobs|/templates|/excel|/docs|/openapi.json|/redoc`
на `127.0.0.1:8811`, `client_max_body_size 200M`, в proxy-location —
`proxy_read_timeout 300s; proxy_send_timeout 300s;` (превью и экспорт ждут
очереди LibreOffice; при нескольких одновременных задачах рендер идёт дольше
стандартных 60 с и без этого клиент получает 504, хотя backend работает).
