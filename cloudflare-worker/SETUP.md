# Public-site Refresh button — Cloudflare Worker setup

This 5-minute setup gives the **public** dashboard (the GitHub Pages site)
a working Refresh button. After this, clicking Refresh on
`https://islamxnajmul.github.io/cardiotrack-dashboard/` triggers a real
cloud sync — same as if you'd clicked "Run workflow" in GitHub Actions —
and the page auto-reloads when the new build is ready, ~60-90 seconds later.

The Worker is a 90-line proxy that holds your GitHub Personal Access Token
securely (it can't be embedded in the browser without exposing it). Cloudflare
Workers free tier handles 100,000 requests/day, far more than this needs.

---

## What you'll do

1. Get a **fine-grained GitHub PAT** with workflow-trigger permission (you may
   already have one from earlier setup).
2. Install Cloudflare's `wrangler` CLI on your Mac.
3. Deploy the Worker (one command).
4. Set the PAT as a Worker secret (one command).
5. Paste the Worker URL into your local config so the dashboard knows where
   to send Refresh requests.

---

## Step 1 — Get a fine-grained GitHub PAT

If you already have a working PAT with `workflow` scope from the original
hosting setup, you can reuse it. Otherwise:

1. Open <https://github.com/settings/tokens?type=beta>
2. **Generate new token (fine-grained)**.
3. Name: `cardiotrack-sync-worker`
4. Expiration: 1 year (longest allowed — fine for personal use)
5. **Repository access**: Only select repositories → pick `cardiotrack-dashboard`
6. **Repository permissions** — set:
   - **Actions**: Read and write
   - **Contents**: Read-only
   - **Metadata**: Read-only (auto-required)
7. **Generate token** and **copy it** (starts with `github_pat_…`). You only see it once.

---

## Step 2 — Install wrangler CLI

In your Mac Terminal:

```bash
brew install cloudflare-wrangler
```

If you don't have Homebrew, see <https://brew.sh>.

Then sign in:

```bash
wrangler login
```

A browser tab opens; click **Allow** to grant CLI access to your Cloudflare
account.

---

## Step 3 — Deploy the Worker

```bash
cd "/Users/najmulislam/Documents/Claude/Projects/Sales/Cardiotrack sales/cloudflare-worker"
wrangler deploy
```

This prints something like:

```
Total Upload: 2.36 KiB
Uploaded cardiotrack-sync-trigger (1.2 sec)
Deployed cardiotrack-sync-trigger triggers (3.1 sec)
  https://cardiotrack-sync-trigger.<your-subdomain>.workers.dev
```

**Copy that URL** — you'll need it in Step 5.

---

## Step 4 — Add the PAT as a Worker secret

```bash
cd "/Users/najmulislam/Documents/Claude/Projects/Sales/Cardiotrack sales/cloudflare-worker"
wrangler secret put GH_TOKEN
```

When prompted, paste the PAT you generated in Step 1. Press Enter.

You can verify it took:

```bash
wrangler secret list
```

Should print something like:
```
[{"name":"GH_TOKEN","type":"secret_text"}]
```

---

## Step 5 — Tell the dashboard where the Worker is

On your Mac, edit `Automation/local_config.json` (create it from
`local_config.example.json` if it doesn't exist):

```json
{
  "quarter_target_file_id": "...your-existing-id...",
  "public_refresh_url": "https://cardiotrack-sync-trigger.<your-subdomain>.workers.dev"
}
```

Then rebuild and push so the Worker URL is baked into the deployed
dashboard:

```bash
cd "/Users/najmulislam/Documents/Claude/Projects/Sales/Cardiotrack sales"
python3 Automation/cardiotrack_sync.py --rebuild
git add Data/Output/
git commit -m "Configure public refresh Worker URL"
git push
```

---

## Step 6 — Test

1. Wait ~60 seconds for GitHub Pages to redeploy the updated HTML.
2. Open <https://islamxnajmul.github.io/cardiotrack-dashboard/> (hard-refresh: ⌘-Shift-R).
3. Click **⟳ Refresh** in the dashboard header.
4. You should see a modal: *"Sync in progress… Pulling Drive + Gmail …"*
5. The modal polls the workflow's status. When it finishes (~60s), the page
   auto-reloads with the fresh data.

---

## Troubleshooting

| Issue | Fix |
|---|---|
| `wrangler: command not found` | Reinstall: `brew install cloudflare-wrangler`. Or use `npm install -g wrangler` if you prefer Node. |
| Worker returns `500: Worker missing GH_TOKEN` | Step 4 was skipped or the secret didn't save. Re-run `wrangler secret put GH_TOKEN`. |
| Refresh button still shows "static snapshot" modal | `local_config.json` doesn't have `public_refresh_url`, OR you didn't rebuild + push after editing. Rerun Step 5. |
| `403 Forbidden` from GitHub when Worker tries to dispatch | PAT lacks `Actions: Read and write` permission. Edit the token, add it, save. The same PAT string keeps working with the new permissions. |
| Worker is reachable but `POST /` returns `204` and nothing happens in GitHub | That actually IS success — GitHub returns 204 on a successful dispatch. Check the Actions tab; a run should appear within seconds. |

---

## What the Worker actually does

```
                ┌─────────────────────────────────────┐
                │  Public dashboard (github.io)       │
                │  user clicks ⟳ Refresh              │
                └────────────────┬────────────────────┘
                                 │  POST https://…workers.dev/
                                 ▼
                ┌─────────────────────────────────────┐
                │  Cloudflare Worker (this code)      │
                │  - reads GH_TOKEN from secret       │
                │  - calls GitHub workflow_dispatch   │
                └────────────────┬────────────────────┘
                                 │  POST .../actions/workflows/sync.yml/dispatches
                                 ▼
                ┌─────────────────────────────────────┐
                │  GitHub Actions                     │
                │  pulls Drive + Gmail + rebuilds     │
                └────────────────┬────────────────────┘
                                 │  push commit
                                 ▼
                ┌─────────────────────────────────────┐
                │  GitHub Pages auto-redeploy         │
                └─────────────────────────────────────┘
                                 │  poll /status until run completes
                                 ▼
                       Dashboard auto-reloads
```

Total elapsed: ~60-90 seconds. Worker uses ~5 requests per sync, so even at
20 syncs per day you're at 100 / 100,000 free tier budget.
