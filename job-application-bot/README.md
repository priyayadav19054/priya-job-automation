# Job Application Bot — V1

Safe first version: reads one PENDING job, opens its URL in visible Chromium, and takes a screenshot. It does not fill or submit anything yet.

## Setup
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
pdflatex --version
```

Put your workbook at `input/job-tracker.xlsx` and run:
```bash
python apply_jobs.py
```
