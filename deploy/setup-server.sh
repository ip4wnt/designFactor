#!/usr/bin/env bash
# Первичная установка DesignFactor на ЧИСТЫЙ Ubuntu 24.04 (нужен Python >= 3.11).
#
# Что делает: ставит пакеты (LibreOffice, poppler, nginx, earlyoom), создаёт
# системного пользователя, venv и зависимости, .env, systemd-сервис с лимитами
# памяти под объём RAM этой машины, swap, earlyoom, самопроверку и nginx.
#
# Запуск (код уже распакован в /opt/designfactor, см. deploy/README.md):
#   sudo bash /opt/designfactor/deploy/setup-server.sh
# Повторный запуск безопасен: существующие .env, swap и пользователь не трогаются.
set -euo pipefail

APP_USER=designfactor
APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PORT=8811

log() { printf '\n==> %s\n' "$*"; }
die() { printf 'ОШИБКА: %s\n' "$*" >&2; exit 1; }

[ "$(id -u)" -eq 0 ] || die "запустите от root: sudo bash $0"
[ -f "$APP_DIR/backend/app/main.py" ] || die "не найден backend/app/main.py в $APP_DIR — распакуйте проект целиком"
case "$APP_DIR" in /root/*|/home/*) die "распакуйте проект в /opt/designfactor: пользователю сервиса нет доступа в $APP_DIR" ;; esac
cd "$APP_DIR"  # иначе sudo -u наследует недоступный пользователю сервиса каталог (например /root)

# --- ресурсы машины -> лимиты (правила из deploy/README.md) -----------------
RAM_MB=$(awk '/MemTotal/{print int($2/1024)}' /proc/meminfo)
MEM_HIGH_MB=$((RAM_MB * 60 / 100))
MEM_MAX_MB=$((RAM_MB * 75 / 100))
JOBS=$((RAM_MB / 3072)); [ "$JOBS" -ge 1 ] || JOBS=1
SWAP_MB=$((RAM_MB / 4)); [ "$SWAP_MB" -ge 1024 ] || SWAP_MB=1024
log "RAM ${RAM_MB} МБ -> MemoryHigh ${MEM_HIGH_MB}M, MemoryMax ${MEM_MAX_MB}M, JOB_CONCURRENCY ${JOBS}, swap ${SWAP_MB}M"
[ "$RAM_MB" -ge 3500 ] || echo "ВНИМАНИЕ: меньше 4 ГБ RAM — генерация с LibreOffice может упираться в память."

# --- пакеты -----------------------------------------------------------------
log "Установка пакетов"
export DEBIAN_FRONTEND=noninteractive
apt-get update -y
apt-get install -y python3 python3-venv python3-pip libreoffice-impress poppler-utils \
    fonts-liberation fonts-dejavu-core fonts-crosextra-carlito nginx earlyoom curl

python3 - <<'PY' || die "нужен Python >= 3.11 (Ubuntu 24.04), у вас: $(python3 --version)"
import sys
sys.exit(0 if sys.version_info >= (3, 11) else 1)
PY

# --- пользователь и права ---------------------------------------------------
log "Пользователь $APP_USER"
id "$APP_USER" >/dev/null 2>&1 || useradd --system --create-home --home-dir "/var/lib/$APP_USER" \
    --shell /usr/sbin/nologin "$APP_USER"
chown -R "$APP_USER:$APP_USER" "$APP_DIR"

# --- venv и зависимости -----------------------------------------------------
log "Python-окружение и зависимости (несколько минут)"
sudo -u "$APP_USER" python3 -m venv "$APP_DIR/backend/venv"
sudo -u "$APP_USER" "$APP_DIR/backend/venv/bin/pip" install --upgrade pip
sudo -u "$APP_USER" "$APP_DIR/backend/venv/bin/pip" install -r "$APP_DIR/backend/requirements.txt"

# --- .env -------------------------------------------------------------------
ENV_FILE="$APP_DIR/backend/.env"
set_env() {
    if grep -q "^$1=" "$ENV_FILE"; then sed -i "s|^$1=.*|$1=$2|" "$ENV_FILE"; else echo "$1=$2" >> "$ENV_FILE"; fi
}
if [ ! -f "$ENV_FILE" ]; then
    log "Создаю backend/.env (ключи впишите сами!)"
    cp "$APP_DIR/backend/.env.example" "$ENV_FILE"
    set_env LLM_BASE_URL https://api.alltokens.ru/api/v1
    set_env LLM_MODEL_NAME qwen/qwen3.8-27b
    set_env VLM_BASE_URL https://api.alltokens.ru/api/v1
    set_env VLM_MODEL_NAME qwen/qwen3.8-27b
fi
set_env JOB_CONCURRENCY "$JOBS"
chown "$APP_USER:$APP_USER" "$ENV_FILE"
chmod 600 "$ENV_FILE"

# --- swap -------------------------------------------------------------------
if [ -z "$(swapon --show --noheadings)" ]; then
    log "Swap ${SWAP_MB} МБ"
    fallocate -l "${SWAP_MB}M" /swapfile
    chmod 600 /swapfile
    mkswap /swapfile
    swapon /swapfile
    grep -q '^/swapfile ' /etc/fstab || echo '/swapfile none swap sw 0 0' >> /etc/fstab
    echo 'vm.swappiness=10' > /etc/sysctl.d/90-swappiness.conf
    sysctl --system >/dev/null
else
    log "Swap уже есть — пропускаю"
fi

# --- earlyoom ---------------------------------------------------------------
log "earlyoom"
cp "$APP_DIR/deploy/earlyoom.default" /etc/default/earlyoom
systemctl enable --now earlyoom
systemctl restart earlyoom

# --- systemd-сервис ---------------------------------------------------------
log "Сервис designfactor"
sed -e "s|/home/pankratov/designfactor|$APP_DIR|g" \
    -e "s|^User=pankratov|User=$APP_USER|" \
    -e "s|^MemoryHigh=.*|MemoryHigh=${MEM_HIGH_MB}M|" \
    -e "s|^MemoryMax=.*|MemoryMax=${MEM_MAX_MB}M|" \
    "$APP_DIR/deploy/designfactor.service" > /etc/systemd/system/designfactor.service

# --- самопроверка -----------------------------------------------------------
cp "$APP_DIR/deploy/designfactor-healthcheck.sh" /usr/local/bin/designfactor-healthcheck.sh
chmod +x /usr/local/bin/designfactor-healthcheck.sh
cp "$APP_DIR/deploy/designfactor-healthcheck.service" "$APP_DIR/deploy/designfactor-healthcheck.timer" /etc/systemd/system/

systemctl daemon-reload
systemctl enable designfactor
systemctl restart designfactor
systemctl enable --now designfactor-healthcheck.timer

# --- nginx ------------------------------------------------------------------
log "nginx"
cp "$APP_DIR/deploy/nginx-designfactor.conf" /etc/nginx/sites-available/designfactor
ln -sf /etc/nginx/sites-available/designfactor /etc/nginx/sites-enabled/designfactor
rm -f /etc/nginx/sites-enabled/default
nginx -t
systemctl enable nginx
systemctl reload nginx || systemctl restart nginx

# --- брандмауэр (только если ufw уже включён) -------------------------------
if command -v ufw >/dev/null 2>&1 && ufw status | grep -q "Status: active"; then
    log "ufw: открываю 22, 80, 443"
    ufw allow 22/tcp; ufw allow 80/tcp; ufw allow 443/tcp
fi

# --- проверка ---------------------------------------------------------------
log "Проверка"
sleep 5
code=$(curl -s -m 15 -o /dev/null -w '%{http_code}' "http://127.0.0.1:${PORT}/docs" || echo 000)
echo "backend /docs -> HTTP $code (ожидается 200)"
code=$(curl -s -m 15 -o /dev/null -w '%{http_code}' "http://127.0.0.1/" || echo 000)
echo "nginx   /     -> HTTP $code (ожидается 200)"
command -v soffice >/dev/null && echo "soffice: найден" || echo "ВНИМАНИЕ: soffice не найден — PDF/превью/VLM-аудит не заработают"

cat <<EOF

Готово. Дальше:
  1) впишите ключи:   sudo nano $ENV_FILE      (LLM_API_KEY и VLM_API_KEY)
  2) перезапустите:   sudo systemctl restart designfactor
  3) откройте в браузере:  http://$(hostname -I | awk '{print $1}')/
  Логи сервиса:       sudo journalctl -u designfactor -f
EOF
