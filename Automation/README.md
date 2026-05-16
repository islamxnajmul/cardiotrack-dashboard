# Cardiotrack Sales Dashboard — Automation

End-to-end workflow that (a) pulls three insurer reports from Gmail every day at midnight,
(b) merges them with the local `Quarter Target.xlsx`, (c) regenerates a self-contained HTML
dashboard, and (d) serves it on `http://localhost:5173` with manual Refresh + Sync buttons.

---

## Tech stack & why

| Layer | Choice | Why |
|---|---|---|
| Language | Python 3.10+ | One language for Gmail OAuth, Excel parsing, and the server. Already in place. |
| Gmail | `google-api-python-client` + OAuth2 desktop flow | Read-only scope, refresh tokens cached locally — no service-account headache. |
| Excel | `openpyxl` | Pure-Python, no Excel installed needed. Handles formulas via `data_only=True`. |
| Server | Flask | Minimal footprint (~one file), binds to `127.0.0.1` only. No production concerns. |
| Scheduler | macOS `launchd` | Native, survives reboots, no extra daemon. Fires at 00:00 every day. |
| Dashboard | Single HTML + Chart.js (CDN) | Zero build step. Works offline (data embedded), but auto-uses the live server when present. |
| Storage | Plain files (`Data/Input/*.xlsx`, `Data/Output/dashboard_data.json`) | Easy to version, easy to inspect, easy to back up. |

**Why not Google Apps Script / Looker Studio?**  Lower flexibility for the pipeline math.
**Why not the cloud (Cloud Run)?**  Overkill for a single-user dashboard that only needs to be live
when the user is working. Re-evaluate if multiple users need shared access.

---

## Data flow

```
            ┌─────────────────────────────────────────────────────┐
            │                  GMAIL INBOX                        │
            │   3 emails:  'Check out the "<report>" report'      │
            └───────────────────────┬─────────────────────────────┘
                                    │  every day 00:00 (launchd)
                                    ▼
            ┌─────────────────────────────────────────────────────┐
            │   cardiotrack_sync.py   →   sync_gmail()            │
            │   • subject:"Check out the ..." filename:xlsx       │
            │   • dedupe by message-id  AND  SHA256(attachment)   │
            │   • writes Data/Input/*.xlsx                        │
            └───────────────────────┬─────────────────────────────┘
                                    ▼
            ┌─────────────────────────────────────────────────────┐
            │   build_dashboard_data()                            │
            │   • reads Quarter Target.xlsx  (Plan, Apr 2026,     │
            │     Overall Billing, BIlling done Insurer Wise)     │
            │   • reads the 3 insurer-wise files                  │
            │   • writes Data/Output/dashboard_data.json          │
            └───────────────────────┬─────────────────────────────┘
                                    ▼
            ┌─────────────────────────────────────────────────────┐
            │   rebuild_html()                                    │
            │   • injects JSON into dashboard_template.html       │
            │   • writes Data/Output/Cardiotrack_Dashboard.html   │
            └───────────────────────┬─────────────────────────────┘
                                    ▼
            ┌─────────────────────────────────────────────────────┐
            │   server.py  (Flask, localhost:5173)                │
            │   GET   /          → dashboard HTML                 │
            │   GET   /api/data  → live JSON                      │
            │   POST  /api/refresh  → re-read Excel               │
            │   POST  /api/sync     → full Gmail + Excel rebuild  │
            └───────────────────────┬─────────────────────────────┘
                                    ▼
                              Browser tab
                         (Refresh & Sync buttons)
```

---

## File layout

```
Cardiotrack sales/
├── Automation/
│   ├── cardiotrack_sync.py          ← Gmail + Excel + HTML pipeline (one file)
│   ├── server.py                    ← Flask web server (localhost:5173)
│   ├── dashboard_template.html      ← template with __DASHBOARD_DATA__ placeholder
│   ├── com.cardiotrack.sync.plist   ← launchd job for 00:00 daily
│   ├── credentials.json             ← (you provide) Google OAuth client
│   ├── token.json                   ← created on first OAuth login
│   ├── sync_log.json                ← processed message IDs + attachment hashes
│   ├── sync.log                     ← script log
│   └── launchd.{out,err}.log        ← launchd's stdout/stderr capture
├── Data/
│   ├── Input/
│   │   ├── Quarter Target.xlsx
│   │   ├── Incoming_Order_Count_Insurer_Wise.csv     ← overwritten by Gmail sync
│   │   ├── Revenue_Generated_Insurer_Wise.csv        ← overwritten by Gmail sync
│   │   └── Closed_Case_Count_Insurer_Wise.csv        ← overwritten by Gmail sync
│   └── Output/
│       ├── dashboard_data.json
│       └── Cardiotrack_Dashboard.html
```

---

## One-time setup

```bash
cd "/Users/najmulislam/Documents/Claude/Projects/Sales/Cardiotrack sales"

# 1. Python deps
pip3 install --user flask google-auth-oauthlib google-api-python-client openpyxl

# 2. Gmail OAuth credentials (one-time)
#    Google Cloud Console → APIs & Services → Credentials →
#      Create OAuth client ID → Desktop app → download JSON
#    Save the downloaded file as:
#        Automation/credentials.json

# 3. First run — opens a browser for OAuth consent, then runs the full pipeline.
python3 Automation/cardiotrack_sync.py

# 4. Install the daily launchd job
cp Automation/com.cardiotrack.sync.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.cardiotrack.sync.plist

# 5. Start the local dashboard server
python3 Automation/server.py
# → open http://localhost:5173
```

---

## Daily usage

- **00:00 every day:** launchd runs `cardiotrack_sync.py` automatically. New report emails
  are pulled in, Excel files refreshed, JSON + HTML regenerated.
- **Whenever you edit `Quarter Target.xlsx`:** click **⟳ Refresh** in the dashboard header.
  The server re-reads Excel and rebuilds the view. No restart needed.
- **Manual Gmail pull mid-day:** click **✉ Sync Gmail**. Same as the 00:00 job but on demand.

The "Last synced" badge in the header shows when `meta.generated_at` was written.

---

## How duplicate prevention works

Two layers — either is enough to skip processing:

1. **Message-ID set** (`sync_log.json` → `processed_message_ids`). Gmail message IDs are
   stable; once we open a message, we never re-open it.
2. **Attachment SHA-256** (`sync_log.json` → `attachment_hashes`). Even if a brand-new email
   carries the exact same `.xlsx` bytes as a prior one, we won't double-count it as "new".

Both are persisted across runs, so the launchd job and manual Sync clicks are idempotent.

---

## Changing email subjects or filenames

Edit two lines at the top of `cardiotrack_sync.py`:

```python
EMAIL_MAP = {
    'Incoming Order Count Insurer Wise': INCOMING_FILE,
    'Revenue Generated Insurer Wise':    REVENUE_FILE,
    'Closed Case Count Insurer Wise':    CLOSED_FILE,
}
SUBJECT_PREFIX = 'Check out the "'
SUBJECT_SUFFIX = '" report'
```

The Gmail search query is built as: `subject:"{PREFIX}{name}{SUFFIX}" has:attachment filename:xlsx`.

---

## Troubleshooting

| Symptom | Where to look |
|---|---|
| Daily run didn't fire | `~/Library/LaunchAgents/com.cardiotrack.sync.plist` loaded? `launchctl list \| grep cardiotrack` |
| OAuth token expired | `python3 Automation/cardiotrack_sync.py --auth` — forces a fresh consent screen and writes a new token.json |
| Need to add a new Google scope (e.g. Drive) | Same `--auth` command — the consent screen will list every scope and granted scopes are confirmed in the terminal afterward |
| Subject changed by sender | Edit `SUBJECT_PREFIX` / `SUBJECT_SUFFIX` in `cardiotrack_sync.py` |
| HTML stuck on old data | Hit **⟳ Refresh** (or hard-reload the tab — Cmd-Shift-R) |
| Server won't start | Port 5173 in use → change `PORT` at top of `server.py` |
| Want to test without Gmail | `python3 Automation/cardiotrack_sync.py --rebuild` |
| Gmail finds an email but doesn't download it | `python3 Automation/cardiotrack_sync.py --gmail-debug` — dumps every search hit with subject, From, attachment filenames + MIME types, and the verdict ("WILL DOWNLOAD" vs "skipped") so you can see exactly why it was rejected |
| Dedup state seems poisoned (e.g. you deleted and re-sent the same report) | `python3 Automation/cardiotrack_sync.py --reset-dedup` — clears `sync_log.json` so the next sync re-evaluates every message from scratch |

Logs:
- `Automation/sync.log` — full script output, tail this first
- `Automation/launchd.err.log` — anything launchd's wrapper caught
- `Automation/sync_log.json` — last-sync timestamp + dedupe state

---

## Scaling beyond one user

When you outgrow localhost (multi-user, off-hours access, mobile):

1. Move `server.py` to **Google Cloud Run** (free tier handles this easily).
2. Move OAuth secret + token to **Google Secret Manager**.
3. Replace launchd with **Cloud Scheduler** → Cloud Run job.
4. Store `dashboard_data.json` in **Cloud Storage**; the HTML loads it via a signed URL.
5. Front it with **Cloudflare Access** for company-only auth.

Estimated cost: **<$10/month** for the volume implied here.
