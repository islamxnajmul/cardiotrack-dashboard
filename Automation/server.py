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

import json, logging, sys, os
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


def _load_local_config() -> dict:
    """Read Automation/local_config.json — optional user-tuned settings.

    Example shape:
        { "quarter_target_file_id": "1ABC…XYZ" }
    """
    cfg_path = sync.AUTO / "local_config.json"
    if not cfg_path.exists():
        return {}
    try:
        return json.loads(cfg_path.read_text())
    except Exception as e:
        log.warning(f"local_config.json could not be parsed: {e}")
        return {}


def _pull_quarter_target_from_drive() -> dict:
    """Best-effort Drive pull of Quarter Target.xlsx into Data/Input/.

    Returns a status dict that gets folded into the sync result.
    The file ID is read from (in priority order):
       1. QUARTER_TARGET_FILE_ID env var (matches CI behaviour)
       2. Automation/local_config.json → quarter_target_file_id
    """
    file_id = (os.environ.get("QUARTER_TARGET_FILE_ID") or
               _load_local_config().get("quarter_target_file_id") or "").strip()
    if not file_id:
        return {"skipped": True, "reason": "no QUARTER_TARGET_FILE_ID configured "
                "(set env var or add quarter_target_file_id to Automation/local_config.json)"}

    # Invoke drive_pull.py as a subprocess so we get the same code path as CI.
    import subprocess
    env = {**os.environ, "QUARTER_TARGET_FILE_ID": file_id}
    try:
        r = subprocess.run([sys.executable, str(sync.AUTO / "drive_pull.py")],
                           capture_output=True, text=True, timeout=60, env=env)
        if r.returncode != 0:
            return {"ok": False, "error": (r.stderr or r.stdout or "non-zero exit").strip()[:500]}
        return {"ok": True, "log": (r.stdout or "").strip()[-400:]}
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": "Drive pull timed out after 60s"}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@app.route("/api/sync", methods=["POST"])
def api_sync():
    """Full upstream pull + rebuild.

    Order of operations:
       1. Pull Quarter Target.xlsx from Google Drive (if configured)
       2. Pull latest Gmail CSV attachments  (force=True by default)
       3. Re-parse Excel + CSVs, rebuild JSON + HTML

    User-initiated 'Sync' clicks default to force=True so dedup is bypassed
    and the newest matching email's attachment is downloaded regardless of
    processed_message_ids.

    Returns Drive + Gmail status dicts so the UI can show what came from
    where — and surface partial failures (e.g. "Gmail OK, Drive failed").

    Query params:
      ?force=false    skip force-mode dedup bypass (matches scheduled cron)
      ?skip_drive=1   skip the Drive pull step
    """
    import os as _os
    force = (request.args.get("force", "true").lower() != "false")
    skip_drive = (request.args.get("skip_drive", "0") == "1")

    # ── Step 1: pull Quarter Target from Drive ─────────────────────────
    drive_status = None
    if not skip_drive:
        drive_status = _pull_quarter_target_from_drive()
        if drive_status.get("ok") is False:
            log.error(f"Drive pull failed: {drive_status.get('error')}")
        elif drive_status.get("skipped"):
            log.info(f"Drive pull skipped: {drive_status.get('reason')}")
        else:
            log.info("Drive pull OK ✓")

    # ── Step 2 & 3: Gmail + rebuild (unchanged) ────────────────────────
    try:
        data = sync.run_sync(skip_gmail=False, require_gmail=True, force_gmail=force)
        gmail = data.get("meta", {}).get("gmail_status", {})
        # Stash drive status into meta so the dashboard can show it too
        if drive_status is not None:
            data.setdefault("meta", {})["drive_status"] = drive_status
            (sync.DATA_JSON).write_text(json.dumps(data, indent=2, ensure_ascii=False))
        return jsonify({
            "ok":            True,
            "generated_at":  data["meta"]["generated_at"],
            "force":         force,
            "drive":         drive_status,
            "gmail":         gmail,
        })
    except Exception as e:
        log.exception("sync failed")
        return jsonify({"ok": False, "error": str(e),
                         "drive": drive_status}), 500


@app.route("/api/pull_drive", methods=["POST"])
def api_pull_drive():
    """Pull Quarter Target.xlsx from Drive only — does NOT rebuild.
    Used by the Refresh button when the user just edited the sheet and
    wants the local Excel mirror updated before clicking Refresh."""
    s = _pull_quarter_target_from_drive()
    if s.get("ok") is False:
        return jsonify({"ok": False, "error": s.get("error")}), 500
    return jsonify({"ok": True, "drive": s})


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
