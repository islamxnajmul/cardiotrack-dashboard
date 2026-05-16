#!/bin/bash
# ─────────────────────────────────────────────────────────────────────
# prepare_secrets.sh
# Prints the three GitHub Secrets you need to add to your repo.
# Run from the Cardiotrack sales/ folder:
#     bash Automation/prepare_secrets.sh
# Then copy each block exactly (including newlines) into the matching
# Settings → Secrets → Actions → New repository secret entry.
# ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

print_block() {
  local name="$1" file="$2"
  echo ""
  echo "════════════════════════════════════════════════════════════════"
  echo "  Secret name:  $name"
  echo "════════════════════════════════════════════════════════════════"
  if [ -f "$file" ]; then
    cat "$file"
  else
    echo "(file not found: $file)"
    echo "Run cardiotrack_sync.py once to generate it, then re-run this script."
  fi
  echo ""
}

clear
echo "▶ Copy each of the following blocks into a new GitHub Secret of the same name."
echo "  Repo URL → Settings → Secrets and variables → Actions → New repository secret"

print_block "GOOGLE_CREDENTIALS_JSON" "credentials.json"
print_block "GOOGLE_TOKEN_JSON"        "token.json"

echo "════════════════════════════════════════════════════════════════"
echo "  Secret name:  QUARTER_TARGET_FILE_ID"
echo "════════════════════════════════════════════════════════════════"
echo "  Paste the Drive file ID — the long random string in the share URL"
echo "  between /d/ and /edit . For example, the URL:"
echo "      https://docs.google.com/spreadsheets/d/1ABC…XYZ/edit"
echo "  has file ID:  1ABC…XYZ"
echo "════════════════════════════════════════════════════════════════"
echo ""
echo "Done. Add all three to GitHub, then push the repo — the workflow will run."
