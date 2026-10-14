#!/usr/bin/env bash
# Legt die Lodgify-Webhooks an und speichert die Signatur-Secrets direkt als Worker-Secrets
# (Lodgify gibt das Secret nur bei der Erstellung zurück; es wird nicht ausgegeben).
# Nutzung (im Ordner webhook-relay): LODGIFY_API_KEY=... ./subscribe-lodgify.sh https://<worker>.workers.dev
set -uo pipefail
BASE="${1:?Worker-URL angeben}"
BASE="${BASE%/}"
KEY="$(printf %s "${LODGIFY_API_KEY:?LODGIFY_API_KEY fehlt}" | tr -d '[:space:]')"
fail=0
for ev in booking_new_any_status booking_change booking_status_change_booked \
          booking_status_change_tentative booking_status_change_open booking_status_change_declined; do
  code=$(curl -s -o /tmp/lodgify-sub.out -w "%{http_code}" https://api.lodgify.com/webhooks/v1/subscribe \
    -X POST -H "X-ApiKey: ${KEY}" -H "Content-Type: application/json" \
    -d "{\"event\":\"$ev\",\"target_url\":\"$BASE/lodgify/$ev\"}")
  echo "$ev -> $code"   # 201 = ok, 409 = existiert bereits (Secret dann nicht erneut abrufbar)
  if [ "$code" = "201" ]; then
    secret=$(python3 -c 'import json;print(json.load(open("/tmp/lodgify-sub.out"))["secret"])')
    name="LODGIFY_SECRET_$(printf %s "$ev" | tr '[:lower:]' '[:upper:]')"
    printf %s "$secret" | npx wrangler secret put "$name" >/dev/null && echo "  Secret $name gespeichert" || fail=1
  elif [ "$code" != "409" ]; then
    cat /tmp/lodgify-sub.out; echo; fail=1
  fi
done
rm -f /tmp/lodgify-sub.out
exit $fail
