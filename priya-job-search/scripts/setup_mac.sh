#!/bin/bash
set -e
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
python3 -m playwright install chromium
echo "Setup complete. Run: source .venv/bin/activate && python3 run_all.py"
