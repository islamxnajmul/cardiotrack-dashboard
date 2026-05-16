#!/usr/bin/env python3
"""
Download Quarter Target.xlsx from Google Drive into Data/Input/.

Used by the GitHub Actions workflow so the daily cloud sync can read the file
the user edits locally (auto-synced via the Mac's Google Drive client).

Reuses the same OAuth token.json as cardiotrack_sync.py — that token must have
both gmail.readonly AND drive.readonly scopes. If you previously authorised
only Gmail, delete token.json and run cardiotrack_sync.py once on your Mac;
the OAuth browser flow will reauthorise with the expanded scope.

Required environment variable:
    QUARTER_TARGET_FILE_ID — Drive file ID for Quarter Target.xlsx
                            (the bit in the URL after /d/ and before /edit)

Usage:
    QUARTER_TARGET_FILE_ID=abc123… python3 drive_pull.py
"""
import os
import sys
from pathlib import Path

try:
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
except ImportError:
    print("⚠  Google API packages missing.  pip install google-auth-oauthlib "
          "google-api-python-client", file=sys.stderr)
    sys.exit(1)

AUTO       = Path(__file__).parent
INPUT_DIR  = AUTO.parent / "Data" / "Input"
TOKEN_FILE = AUTO / "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Native xlsx MIME (vs. Google Sheets which is application/vnd.google-apps.spreadsheet).
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def main():
    file_id = os.environ.get("QUARTER_TARGET_FILE_ID", "").strip()
    if not file_id:
        print("ERROR: QUARTER_TARGET_FILE_ID env var is not set.", file=sys.stderr)
        print("       Set it to the Drive file ID — the long random string in", file=sys.stderr)
        print("       the file's share URL between /d/ and /edit .", file=sys.stderr)
        sys.exit(2)

    if not TOKEN_FILE.exists():
        print(f"ERROR: {TOKEN_FILE} not found.", file=sys.stderr)
        print("       In CI: restore it from the GOOGLE_TOKEN_JSON secret first.", file=sys.stderr)
        print("       Locally: run cardiotrack_sync.py once to do the OAuth flow.", file=sys.stderr)
        sys.exit(3)

    creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        TOKEN_FILE.write_text(creds.to_json())

    service = build("drive", "v3", credentials=creds)

    # Determine whether this is a native Excel upload or a Google Sheet —
    # the download API call is different for each.
    meta = service.files().get(fileId=file_id, fields="name,mimeType").execute()
    src_mime = meta.get("mimeType", "")
    print(f"  Source file: {meta.get('name')!r}  ({src_mime})")

    if src_mime == "application/vnd.google-apps.spreadsheet":
        request = service.files().export_media(fileId=file_id, mimeType=XLSX_MIME)
    else:
        request = service.files().get_media(fileId=file_id)

    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = INPUT_DIR / "Quarter Target.xlsx"

    with open(out_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"  Downloading… {int(status.progress() * 100)}%")
    print(f"✓ Saved {out_path}  ({out_path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
