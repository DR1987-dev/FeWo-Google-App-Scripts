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
