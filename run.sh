#!/usr/bin/env bash
# Quick-start runner for the Costing Desk app.
set -e
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "▶ Creating virtual environment…"
  python3 -m venv .venv
fi

# shellcheck source=/dev/null
source .venv/bin/activate

echo "▶ Installing dependencies…"
pip install -q --upgrade pip
pip install -q -r requirements.txt

echo "▶ Starting Flask server on http://127.0.0.1:5000 …"
python app.py
