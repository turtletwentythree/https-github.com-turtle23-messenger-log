# messenger-sync-trigger (Cloudflare Worker)

Public trigger endpoint for the "ð à¸à¸´à¸à¸à¹à¸à¹à¸­à¸¡à¸¹à¸¥à¸¥à¹à¸²à¸ªà¸¸à¸" button on the Messenger
log site. It rate-limits requests via Workers KV and then calls GitHub's
`workflow_dispatch` API to kick off `sync.yml` immediately, so the button
causes a genuinely fresh SharePoint pull instead of just re-fetching the
last scheduled snapshot.

Live at: https://messenger-sync-trigger.turtle23-ai.workers.dev

## Deploys itself

Deployed automatically by `.github/workflows/deploy-worker.yml` whenever
this folder changes on `main` (or via manual workflow_dispatch). That
workflow uses two repo secrets:

- `CF_API_TOKEN` - Cloudflare API token, used by wrangler to deploy.
- `GH_DISPATCH_TOKEN` - a GitHub PAT with `repo`/`workflow` scope, pushed
  into the Worker as the `GH_PAT` secret (via wrangler) so the Worker can
  call GitHub's workflow_dispatch API. Never committed to this repo.

The KV namespace binding (`SYNC_STATE`) is declared in `wrangler.toml` -
its id is not sensitive, so it's fine to commit.

## Local reference (not required for deploys)

```
cd cloudflare-worker
npx wrangler deploy
npx wrangler secret put GH_PAT
```
