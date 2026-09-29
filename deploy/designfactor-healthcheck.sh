#!/usr/bin/env bash
# Самопроверка backend: раз в минуту (designfactor-healthcheck.timer) дёргает
# /docs с таймаутом 10 с. Три неответа подряд -> systemctl restart. В журнал
# каждый раз пишется свободная память — по ней после инцидента видно динамику.
set -u
STATE=/run/designfactor-healthcheck.fails
URL=http://127.0.0.1:8811/docs
code=$(curl -s -m 10 -o /dev/null -w '%{http_code}' "$URL" || echo 000)
mem=$(free -m | awk '/^Mem:/{printf "used=%dM avail=%dM", $3, $7} /^Swap:/{printf " swap_used=%dM", $3}')
if [ "$code" = "200" ]; then
  echo 0 > "$STATE"
  echo "ok http=$code $mem"
  exit 0
fi
fails=$(( $(cat "$STATE" 2>/dev/null || echo 0) + 1 ))
echo "$fails" > "$STATE"
echo "FAIL http=$code fails=$fails $mem"
if [ "$fails" -ge 3 ]; then
  echo "три неответа подряд — перезапуск designfactor.service"
  systemctl restart designfactor.service
  echo 0 > "$STATE"
fi
