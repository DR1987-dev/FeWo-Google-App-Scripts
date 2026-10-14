#!/usr/bin/env bash
# Legt die Lexware-Event-Subscriptions an.
# Nutzung: LEXWARE_API_KEY=... ./subscribe.sh https://<worker>.workers.dev
# Hinweis: Subscriptions werden gelöscht, wenn der API-Key widerrufen wird -> neu ausführen.
set -uo pipefail
CALLBACK="${1:?Callback-URL angeben}"
[ -n "${LEXWARE_API_KEY:-}" ] || { echo "LEXWARE_API_KEY ist nicht gesetzt." >&2; exit 1; }
KEY="$(printf %s "$LEXWARE_API_KEY" | tr -d '[:space:]')"
fail=0
for t in voucher.created voucher.changed voucher.deleted voucher.status.changed \
         invoice.created invoice.changed invoice.deleted invoice.status.changed \
         credit-note.created credit-note.changed credit-note.deleted \
         payment.changed contact.created contact.changed contact.deleted; do
  code=$(curl -s -o /tmp/sub.out -w "%{http_code}" https://api.lexware.io/v1/event-subscriptions \
    -X POST -H "Authorization: Bearer ${KEY}" -H "Content-Type: application/json" \
    -d "{\"eventType\":\"$t\",\"callbackUrl\":\"$CALLBACK\"}")
  echo "$t -> $code"   # 201 = ok, 409 = existiert bereits
  if [ "$code" != "201" ] && [ "$code" != "409" ]; then cat /tmp/sub.out; echo; fail=1; fi
  sleep 1
done
exit $fail
