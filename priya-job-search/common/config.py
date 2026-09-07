import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

def load_config():
    with open(ROOT / "config" / "job_search.json", "r", encoding="utf-8") as f:
        return json.load(f)
