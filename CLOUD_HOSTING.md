# Cloud Hosting — GitHub Actions + Cloudflare Pages

This document is the step-by-step setup to take your local Cardiotrack
dashboard and host it on the public internet, fully free, with zero ongoing
maintenance.

**End result:** every morning at 00:00 UTC, GitHub Actions pulls fresh CSV
attachments from Gmail and your latest Quarter Target.xlsx from Google Drive,
regenerates the dashboard, and pushes the updated HTML to Cloudflare Pages.
Your team opens `https://cardiotrack.pages.dev` (or your custom domain) and
sees the latest data — no Macs need to be awake.

---

## Architecture recap

```
                 Daily 00:00 UTC
                       │
                       ▼
    ┌──────────────────────────────────────────┐
    │  GitHub Actions  (.github/workflows/     │
    │                   sync.yml)              │
    │                                          │
    │  1. Pulls Quarter Target.xlsx from your  │
    │     Google Drive (auto-synced from Mac)  │
    │  2. Runs cardiotrack_sync.py             │
    │     ├─ Gmail → 3 CSVs                    │
    │     └─ Excel + CSVs → dashboard_data.json│
    │  3. Commits regenerated HTML/JSON to     │
    │     the repo                             │
    └──────────────────┬───────────────────────┘
                       │ git push
                       ▼
    ┌──────────────────────────────────────────┐
    │  GitHub repo  (private — your data       │
    │                stays private)            │
    └──────────────────┬───────────────────────┘
                       │ push hook
                       ▼
    ┌──────────────────────────────────────────┐
    │  Cloudflare Pages (free, unlimited       │
    │  bandwidth)                              │
    │  ──→  https://cardiotrack.pages.dev      │
    └──────────────────┬───────────────────────┘
                       ▼
                  Team browsers
```

Total ongoing cost: **₹0 / month**.

---

## One-time setup — 45 minutes

### Step 1 — Add Google Drive scope to your OAuth (5 min)

Locally on your Mac, the script's OAuth token needs Drive read permission
(it didn't before). Run the dedicated `--auth` mode — it forces a fresh
consent screen showing both scopes:

```bash
cd "/Users/najmulislam/Documents/Claude/Projects/Sales/Cardiotrack sales"
python3 Automation/cardiotrack_sync.py --auth
```

What to expect:

1. The terminal prints a banner reading *"Browser-based Google sign-in
   required. A browser tab will open in a moment."*
2. Your default browser opens to Google's account picker — pick your Google
   account.
3. The next screen says *"<app name> wants to access your Google Account"*
   and lists **two** permissions:
   - **Read your email messages and settings**
   - **See and download all your Google Drive files**
4. Click **Allow** (or "Continue" → "Allow"). If you only see Gmail or only
   see Drive, scroll down — both must be checked / approved.
5. The browser tab will say *"The authentication flow has completed. You may
   close this window."*
6. Back in the terminal you'll see:
   ```
   ✓ Authenticated as: you@example.com
   ✓ Token saved to:   .../Automation/token.json
   ✓ Scopes granted:
      • https://www.googleapis.com/auth/gmail.readonly
      • https://www.googleapis.com/auth/drive.readonly
   ```

That confirms both scopes are on the new token. If you only see one scope
listed, redo Step 1 and approve every permission prompt this time.

**If the browser doesn't open at all:**
- The script falls back to printing a URL to the terminal — copy-paste it
  into any browser to continue.
- Check that `Automation/credentials.json` exists (OAuth client JSON
  downloaded from Google Cloud Console — see Automation/README.md if you
  haven't done this yet).
- If the script errors with *"redirect_uri_mismatch"*, your OAuth client is
  the wrong type — must be **Desktop app** in Google Cloud Console, not Web.

### Step 2 — Put Quarter Target.xlsx in Google Drive (5 min)

You're probably already running Google Drive on your Mac. Move
`Quarter Target.xlsx` into your Drive folder (or anywhere inside Google Drive
in Finder). The Drive client auto-syncs every edit you make in Excel.

Once it's in Drive, get the file's ID:

1. Open drive.google.com in a browser.
2. Right-click the file → **Get link** → **Copy link**.
3. The URL looks like:
   `https://docs.google.com/spreadsheets/d/1ABCdef…XYZ/edit?usp=sharing`
4. Copy the chunk between `/d/` and `/edit` — that's the file ID. Save it.

### Step 3 — Create a private GitHub repo (5 min)

1. Go to [github.com/new](https://github.com/new).
2. Repo name: `cardiotrack-dashboard` (or anything).
3. **Private**. (Your data stays private; only the GH Actions runner sees it.)
4. Don't initialise with anything — push from local.
5. In Terminal:

```bash
cd "/Users/najmulislam/Documents/Claude/Projects/Sales/Cardiotrack sales"
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<YOUR_USERNAME>/cardiotrack-dashboard.git
git push -u origin main
```

The `.gitignore` keeps `credentials.json`, `token.json`, `Quarter Target.xlsx`
and the daily CSVs out of the repo automatically.

**If you've run this step before and hit errors:**

| Error | Fix |
|---|---|
| `fatal: remote origin already exists.` | Use `git remote set-url origin <URL>` instead of `git remote add origin <URL>` — it updates the existing remote in place. Check the current value first with `git remote -v`. |
| `nothing to commit, working tree clean` | The commit was already made on a previous run. Skip straight to `git push -u origin main`. |
| `error: failed to push some refs` / `Updates were rejected` | The remote already has commits (e.g. you initialised the repo with a README on GitHub). Run `git pull --rebase origin main` then `git push -u origin main`. |
| `remote: Invalid username or token. Password authentication is not supported for Git operations.` | GitHub disabled password auth in 2021. **Easiest fix:** install [GitHub CLI](https://cli.github.com/) — `brew install gh` then `gh auth login` (pick GitHub.com → HTTPS → Yes → Login with web browser). Retry the push and it works. **Alternative:** create a Personal Access Token at https://github.com/settings/tokens?type=beta (Contents: Read+write on this repo), paste it as the "password" when git prompts, then run `git config --global credential.helper osxkeychain` so you only need to paste it once. |
| `Permission denied (publickey)` | You're using SSH but no SSH key is set up. Easiest: switch to HTTPS with `git remote set-url origin https://github.com/<USER>/<REPO>.git` and follow the row above. |
| `On branch master` instead of `main` | Older Git default. Run `git branch -M main` first. |

You can re-run any of these commands safely — none of them destroy local work.

### Step 4 — Add three repo secrets (5 min)

Generate the secret values with the helper:

```bash
bash Automation/prepare_secrets.sh
```

It prints three blocks. For each, on GitHub:

> **Settings → Secrets and variables → Actions → New repository secret**

| Secret name | Value |
|---|---|
| `GOOGLE_CREDENTIALS_JSON` | The full JSON contents of `credentials.json` |
| `GOOGLE_TOKEN_JSON` | The full JSON contents of `token.json` |
| `QUARTER_TARGET_FILE_ID` | The file ID you copied in Step 2 |

### Step 5 — Test the workflow (5 min)

In the GitHub repo:

1. **Actions** tab → **Daily Cardiotrack Sync** workflow → **Run workflow** button.
2. Click the running job to watch it. Should finish in ~30-90 seconds.
3. Open the **Workflow summary** at the bottom — it prints the freshly-built
   KPI numbers as a sanity check.
4. If green, the workflow committed updated `Data/Output/*` back to your repo.

If it errors:
- "credentials.json: 0 bytes" → secret value wasn't pasted correctly. Re-add.
- "Drive file not found" → check `QUARTER_TARGET_FILE_ID` value.
- "insufficient permission" → you didn't approve Drive scope in Step 1.
  Delete `token.json`, re-run Step 1, update the `GOOGLE_TOKEN_JSON` secret.

### Step 6 — Connect Cloudflare Pages (10 min)

1. Sign up at [cloudflare.com](https://dash.cloudflare.com/sign-up). Free.
2. In the sidebar, **Workers & Pages** → **Create application** → **Pages** →
   **Connect to Git**.
3. Authorise Cloudflare to read your GitHub repos. Select the
   `cardiotrack-dashboard` repo.
4. **Set up builds and deployments**:
   - **Project name:** `cardiotrack` (you'll get `cardiotrack.pages.dev`).
   - **Production branch:** `main`.
   - **Build command:** *leave empty*.
   - **Build output directory:** `Data/Output`.
   - **Root directory:** *leave blank*.
5. **Save and Deploy**. First deploy takes ~30 seconds.

Once it's done, open `https://cardiotrack.pages.dev` — your dashboard loads.

### Step 7 — (Optional) Custom domain (5 min)

In your Cloudflare Pages project: **Custom domains → Set up a custom
domain**. Point your domain at the Pages app. TLS is automatic and free.

---

## What happens day-to-day

| Action | What happens |
|---|---|
| You edit `Quarter Target.xlsx` in Excel | Google Drive client on your Mac syncs the change within seconds |
| Sales team sends the 3 report emails | They sit in your Gmail inbox |
| 00:00 UTC daily | GH Actions runs: pulls Drive file, pulls Gmail CSVs, regenerates HTML, pushes commit. Cloudflare auto-deploys. Team URL is up-to-date in ~2 minutes |
| You want an out-of-cycle refresh | GitHub → Actions → "Run workflow" button. ~30s end-to-end |
| Daily CSV unchanged | The sync still runs, but if nothing changed, the commit step skips (no-op). Cloudflare doesn't redeploy |

---

## Operating it

| Need to… | Do this |
|---|---|
| Trigger a refresh now | GitHub repo → Actions → run workflow manually |
| See what last ran | GitHub repo → Actions → click the latest run |
| See the actual data the dashboard is serving | `https://cardiotrack.pages.dev/dashboard_data.json` |
| Change the cron time | Edit `cron: '0 0 * * *'` in `.github/workflows/sync.yml` (UTC). `'30 0 * * *'` would be 06:00 IST |
| Add a new team viewer | They just need the URL — Cloudflare Pages is public by default |
| Restrict viewers | Pages → your project → **Access** → Cloudflare Access policy (free for up to 50 users) |
| Token expired | Refresh tokens never expire unless revoked. If yours does, redo Step 1 + update `GOOGLE_TOKEN_JSON` secret |
| Move off Cloudflare | The same repo deploys cleanly on Netlify, Vercel, or GitHub Pages — just point them at `Data/Output` |

---

## Keeping the local Mac setup too

The cloud setup and your existing `Start Dashboard.command` workflow can
coexist — they share the same code. Use whichever is more convenient on a
given day:

- **Local Mac (instant Refresh button)** — when you're actively iterating on
  Quarter Target.xlsx and want to see the impact immediately.
- **Cloud URL (always-on, shareable)** — for the team, and for yourself when
  your Mac is asleep.

The cloud one reflects the previous nightly snapshot (or whatever you
manually triggered last) — up to 24h lag on Excel edits unless you hit the
"Run workflow" button.

---

## Cost & limits

| Resource | Used per month (estimated) | Free-tier limit |
|---|---|---|
| GitHub Actions minutes | ~90 min (3 min × 30 days) | 2,000 min (private repos) |
| Cloudflare Pages builds | ~30 builds | 500 builds |
| Cloudflare Pages bandwidth | depends on traffic | **Unlimited** |
| Gmail API quota | ~30 message reads | 1 billion units/day |
| Drive API quota | ~30 file reads | 1 billion queries/day |

You will not hit any of these limits.
