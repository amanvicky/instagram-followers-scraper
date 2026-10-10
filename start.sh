#!/usr/bin/env bash
# One-click launcher for the Instagram Follower Comparator (macOS/Linux)
cd "$(dirname "$0")"
if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 was not found. Install Python 3.9+ first."
  exit 1
fi
# A venv folder can exist but be broken (e.g. after a Python update).
if [ -x ".venv/bin/python" ] && ! .venv/bin/python --version >/dev/null 2>&1; then
  echo "The existing .venv is broken - rebuilding it..."
  rm -rf .venv
fi
if [ ! -f ".venv/bin/python" ]; then
  echo "First run - setting things up, this takes a couple of minutes..."
  python3 -m venv .venv || exit 1
  .venv/bin/python -m pip install --upgrade pip || exit 1
  .venv/bin/python -m pip install -r requirements.txt || exit 1
  .venv/bin/python -m playwright install chromium || exit 1
  echo "Setup complete."
fi
echo "Starting the app - your browser will open in a moment."
.venv/bin/python app.py
echo "The app has stopped. If an error is printed above, paste it for support."
