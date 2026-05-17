#!/usr/bin/env python3
"""
Cardiotrack Dashboard — local web server
========================================
A tiny Flask app that serves the dashboard HTML and exposes two refresh
endpoints. Binds to 127.0.0.1 only — never reachable from the network.

Endpoints
---------
GET   /                 → Cardiotrack_Dashboard.html
GET   /dashboard_data.json
                        → current JSON (whatever the last sync wrote)
GET   /api/data         → current JSON + meta (for the dashboard's auto-fetch)
POST  /api/refresh      → re-read Excel files + rebuild JSON/HTML.
                          Does NOT hit Gmail. Use this for "Excel changed".
POST  /api/sync         → full Gmail sync + Excel rebuild.

Run
---
    python3 server.py
    # then open http://localhost:5173 in any browser

Requires:  pip3 install flask
"""

import json, logging, sys
from pathlib import Path
from datetime import datetime, timezone

try:
    from flask import Flask, jsonify, send_from_directory, request, abort
except ImportError:
    print("⚠  Flask not installed.  Run:  pip3 install flask")
    sys.exit(1)

# Import the sync module — gives us run_sync(), DATA_JSON, OUTPUT, etc.
sys.path.insert(0, str(Path(__file__).parent))
import cardiotrack_sync as sync

PORT = 5173
HOST = "127.0.0.1"          # localhost only — never bind to 0.0.0.0

app = Flask(__name__, static_folder=None)
log = logging.getLogger("cardiotrack.server")
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-8s  %(message)s",
                    datefmt="%H:%M:%S")


# ─── Routes ──────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    """Serve the dashboard HTML. If it doesn't exist yet, force a rebuild first."""
    if not sync.DASHBOARD_HTML.exists():
        log.info("Dashboard HTML missing — building from scratch")
        sync.run_sync(skip_gmail=True)
    return send_from_directory(str(sync.OUTPUT), "Cardiotrack_Dashboard.html")


@app.route("/dashboard_data.json")
def data_file():
    """Serve the raw JSON the HTML fetches via its built-in loader."""
    if not sync.DATA_JSON.exists():
        sync.run_sync(skip_gmail=True)
    # send_from_directory adds caching headers; we want fresh-every-request.
    resp = send_from_directory(str(sync.OUTPUT), "dashboard_data.json")
    resp.headers["Cache-Control"] = "no-store, must-revalidate"
    return resp


@app.route("/api/data")
def api_data():
    """JSON endpoint with a little extra meta the UI can lean on."""
    if not sync.DATA_JSON.exists():
        sync.run_sync(skip_gmail=True)
    data = json.loads(sync.DATA_JSON.read_text(encoding="utf-8"))
    data.setdefault("meta", {})["served_at"] = datetime.now(timezone.utc).isoformat()
    return jsonify(data)


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    """Re-read Excel files + rebuild JSON & HTML. Skips Gmail."""
    try:
        data = sync.run_sync(skip_gmail=True)
        return jsonify({"ok": True, "generated_at": data["meta"]["generated_at"]})
    except Exception as e:
        log.exception("refresh failed")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sync", methods=["POST"])
def api_sync():
    """Full pipeline: Gmail → Excel → JSON → HTML.

    User-initiated 'Sync Gmail' clicks default to force=True — the user's
    explicit intent is "give me the freshest data", so dedup is bypassed and
    the newest matching email's attachment is downloaded regardless of
    processed_message_ids.

    Returns the full Gmail status dict so the UI can show "downloaded X /
    unchanged Y / missing Z" — not just a green check.
    """
    # `?force=false` opts out (matches the scheduled-cron behaviour). Defaults true.
    force = (request.args.get("force", "true").lower() != "false")
    try:
        data = sync.run_sync(skip_gmail=False, require_gmail=True, force_gmail=force)
        gmail = data.get("meta", {}).get("gmail_status", {})
        return jsonify({
            "ok":            True,
            "generated_at":  data["meta"]["generated_at"],
            "force":         force,
            "gmail":         gmail,
        })
    except Exception as e:
        log.exception("sync failed")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/csv_status")
def api_csv_status():
    """Freshness diagnostic — what CSV files are on disk, how old, and what
    each says about May/Jun 2026 revenue. The dashboard polls this on load
    to surface stale-CSV warnings BEFORE the user wonders why numbers look off.
    """
    out = {"ok": True, "files": {}, "totals_by_month_2026": {}}
    for label, path in (("revenue", sync.REVENUE_FILE),
                        ("incoming", sync.INCOMING_FILE),
                        ("closed",   sync.CLOSED_FILE)):
        if path.exists():
            mt = datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc)
            out["files"][label] = {
                "name":   path.name,
                "mtime":  mt.isoformat(),
                "size":   path.stat().st_size,
                "age_hours": round((datetime.now(timezone.utc) - mt).total_seconds()/3600, 1),
            }
        else:
            out["files"][label] = {"name": path.name, "missing": True}

    # Quick re-sum of the revenue CSV for sanity-check display
    try:
        rev = sync.read_revenue_file(sync.REVENUE_FILE)
        for r in rev:
            m = (r.get("month") or "").strip()
            if "2026" in m:
                out["totals_by_month_2026"][m] = out["totals_by_month_2026"].get(m, 0) + r["amount"]
    except Exception as e:
        out["totals_error"] = str(e)

    # Surface Gmail last-sync timestamp too
    sl = sync.load_sync_log() if sync.SYNC_LOG.exists() else {}
    out["last_gmail_sync"] = sl.get("last_sync")
    return jsonify(out)


@app.route("/api/auth_status")
def api_auth_status():
    """Quick health-check for the Gmail OAuth setup. The dashboard polls this
    on load to decide whether to show the 'Gmail not configured' banner."""
    creds_present = sync.CREDS_FILE.exists()
    token_present = sync.TOKEN_FILE.exists()
    return jsonify({
        "credentials_json": creds_present,
        "token_json":       token_present,
        "configured":       creds_present,
        "ready":            creds_present and token_present,
        "creds_path":       str(sync.CREDS_FILE),
        "hint":             None if creds_present else (
            "credentials.json is missing. Get OAuth Desktop credentials from "
            "Google Cloud Console → APIs & Services → Credentials, then save "
            f"the JSON to {sync.CREDS_FILE}"
        ),
    })


@app.route("/api/status")
def api_status():
    """Health-check / 'when did we last sync' badge."""
    sync_log = sync.load_sync_log() if sync.SYNC_LOG.exists() else {}
    return jsonify({
        "ok": True,
        "last_sync":       sync_log.get("last_sync"),
        "processed_count": len(sync_log.get("processed_message_ids", [])),
        "data_mtime":      datetime.fromtimestamp(sync.DATA_JSON.stat().st_mtime, tz=timezone.utc).isoformat()
                            if sync.DATA_JSON.exists() else None,
    })


if __name__ == "__main__":
    log.info(f"Cardiotrack dashboard server → http://{HOST}:{PORT}")
    log.info("Open that URL in a browser. Ctrl-C to stop.")
    # debug=False so Flask doesn't auto-reload mid-rebuild
    app.run(host=HOST, port=PORT, debug=False)
