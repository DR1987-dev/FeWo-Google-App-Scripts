# Lexware Webhook Relay (Cloudflare Worker)

Lexware webhook -> Worker (verifies `X-Lxo-Signature`) -> GitHub `repository_dispatch` -> `importLexwareDelta`.

## Setup

1. Free account at https://dash.cloudflare.com (no credit card).
2. GitHub: Settings > Developer settings > Fine-grained tokens. Select only this repo, permission **Contents: Read and write**.
3. In this folder:

```bash
npx wrangler login
npx wrangler secret put GITHUB_TOKEN      # the token from step 2
npx wrangler secret put GITHUB_REPO       # DR1987-dev/FeWo-Google-App-Scripts
npx wrangler secret put ORGANIZATION_ID   # optional: Lexware organizationId
npx wrangler deploy                       # prints https://lexware-relay.<name>.workers.dev
```

4. Subscribe to the Lexware events (re-run after any API key change):

```bash
LEXWARE_API_KEY=... ./subscribe.sh https://lexware-relay.<name>.workers.dev
```

5. Test: change something in Lexware, then check the Actions tab for a "Run FeWo Skripte" run triggered by `repository_dispatch`.

## Lodgify

Lodgify signiert jeden Webhook (`ms-signature: sha256=<HMAC-SHA256 hex>`) mit einem eigenen Secret je Webhook, das nur bei der Erstellung ausgegeben wird. `subscribe-lodgify.sh` legt die Webhooks an und schreibt die Secrets direkt in den Worker.

```bash
npx wrangler deploy
LODGIFY_API_KEY=... ./subscribe-lodgify.sh https://lexware-relay.<name>.workers.dev
```

- Callback-URLs: `<worker>/lodgify/<event>` (Lodgify verlangt pro Event eine eindeutige URL).
- Bei `409` existiert der Webhook schon, das Secret ist nicht erneut abrufbar. Dann über `DELETE /webhooks/v1/unsubscribe` entfernen und neu anlegen.
- Alle Webhooks lösen den Workflow `runDelta` aus (Lodgify-Import, Lexware-Delta, Rechnungs-Upload, Planung, Monatswerte).
- Worker-Logs live ansehen (im Relay-Verzeichnis): `./tail-lodgify.sh`. Die Diagnose protokolliert nur Event-Typ, erkannte ID-Quelle und Feldnamen – keine Payload-Werte.
