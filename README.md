# Priya Job Search — Local

Local job-search pipeline for LinkedIn + Naukri.

## What this does

- Searches the configured roles across the configured locations.
- Runs LinkedIn and Naukri independently.
- Combines results.
- Deduplicates jobs.
- Applies the experience rule: include jobs whose minimum stated experience requirement is <= 5 years; jobs with no stated requirement are included.
- Ranks jobs.
- Writes an Excel tracker and JSON output.
- Can optionally send an email.
- Can be scheduled locally on macOS later.

## Current searches

Roles:
1. Java Backend Developer
2. Java Developer Spring Boot
3. Java Microservices
4. Spring Boot Backend
5. Java Software Engineer

Locations:
- Pune, Maharashtra, India
- Gurugram, Haryana, India
- Noida, Uttar Pradesh, India
- Bengaluru, Karnataka, India
- Hyderabad, Telangana, India
- Remote India

## Setup

Install Python 3 first if needed.

Then from the repository folder:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
python3 -m playwright install chromium
```

Run the complete local search:

```bash
python3 run_all.py
```

The browsers are visible by default so you can see what is happening.

## Outputs

Generated files go under `jobs/`:

- `jobs/job-tracker.xlsx`
- `jobs/jobs.json`
- `jobs/diagnostics/`

Do not commit credentials or `.env` files.
