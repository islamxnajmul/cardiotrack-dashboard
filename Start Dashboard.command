#!/bin/bash
# ──────────────────────────────────────────────────────────────────────────
# Cardiotrack Dashboard — double-click to start
# ──────────────────────────────────────────────────────────────────────────
# Double-clicking this file opens Terminal, installs Flask if missing, and
# starts the local dashboard server on http://localhost:5173 .
# Leave the Terminal window open while you're using the dashboard.
# Press Ctrl-C (or close the window) to stop.
# ──────────────────────────────────────────────────────────────────────────

# Make the script self-locating so it works regardless of where it lives.
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

clear
echo "╔══════════════════════════════════════════════════════════════════════╗"
echo "║          🫀  Cardiotrack Sales Dashboard — local server              ║"
echo "╚══════════════════════════════════════════════════════════════════════╝"
echo ""
echo "Working folder:  $SCRIPT_DIR"
echo ""

# Check python3 exists
if ! command -v python3 >/dev/null 2>&1; then
  echo "✗ python3 not found on this Mac."
  echo "  Install it from https://www.python.org/downloads/  (or run:  brew install python)"
  echo ""
  read -n 1 -s -r -p "Press any key to close this window..."
  exit 1
fi

# Install Flask if missing
echo "Checking Python dependencies…"
if ! python3 -c "import flask" 2>/dev/null; then
  echo "  Flask not installed — installing now (one-time setup)…"
  python3 -m pip install --user flask openpyxl pandas 2>&1 | tail -3
  echo ""
fi

# Install Google API libs too if missing (needed for Gmail sync)
if ! python3 -c "import googleapiclient" 2>/dev/null; then
  echo "  Installing Google API libraries (needed for Gmail sync)…"
  python3 -m pip install --user google-auth-oauthlib google-api-python-client 2>&1 | tail -3
  echo ""
fi

echo "──────────────────────────────────────────────────────────────────────"
echo "  ✓  Starting server…"
echo "  ▶  Open in your browser:    http://localhost:5173"
echo "  ✋ Press Ctrl-C here to stop, or close this Terminal window."
echo "──────────────────────────────────────────────────────────────────────"
echo ""

# Run the server. If it exits, hold the window open so errors stay visible.
python3 "$SCRIPT_DIR/Automation/server.py"
EXIT_CODE=$?

echo ""
echo "──────────────────────────────────────────────────────────────────────"
if [ $EXIT_CODE -ne 0 ]; then
  echo "✗ Server exited with error code $EXIT_CODE."
  echo "  Check the log lines above for details."
else
  echo "✓ Server stopped."
fi
echo "──────────────────────────────────────────────────────────────────────"
echo ""
read -n 1 -s -r -p "Press any key to close this window..."
