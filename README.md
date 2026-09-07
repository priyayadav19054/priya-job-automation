# Priya Job Automation

An end-to-end local job-search and job-application automation project.

The repository is split into two stages:

1. `priya-job-search` — collects jobs from LinkedIn and Naukri, filters, deduplicates, ranks, and generates the job tracker.
2. `job-application-bot` — consumes the tracker and runs the browser-based application workflow.

## Overall Workflow

```text
LinkedIn + Naukri
       ↓
Job Collection
       ↓
Deduplication
       ↓
Experience Filtering
       ↓
Ranking
       ↓
job-tracker.xlsx + jobs.json
       ↓
Resume Prompt
       ↓
Tailored .tex Resume
       ↓
Resume PDF
       ↓
job-application-bot
       ↓
Visible Browser
       ↓
Application Workflow
```

## Repository Structure

```text
priya-job-automation/
├── priya-job-search/
│   └── README.md
├── job-application-bot/
│   └── README.md
└── README.md
```

## Prerequisites

- Python 3.10+
- Git
- Playwright
- Chromium
- Required accounts/credentials for the configured job sources

## Stage 1: Job Search

```bash
cd priya-job-search
```

Follow `priya-job-search/README.md` for setup and execution.

The search pipeline generates:

```text
jobs/job-tracker.xlsx
jobs/jobs.json
```

The Excel tracker contains job details, a clickable job URL, and a job-specific `Resume Prompt`.

## Stage 2: Resume Generation

Use the `Resume Prompt` from the appropriate job row to generate a tailored `.tex` resume.

The prompt provides the original resume and the job description and requires the generated resume to remain:

- Truthful
- ATS-friendly
- One page
- Based only on the original resume
- Tailored to the specific job

Compile the `.tex` file into a one-page PDF.

## Stage 3: Job Application Bot

```bash
cd job-application-bot
```

Follow `job-application-bot/README.md` for setup and execution.

The current V1 workflow opens the job URL in a visible Chromium browser and captures the state for the application workflow.

## Safety

The application bot does not bypass CAPTCHA, OTP, authentication challenges, or other anti-automation/security mechanisms.

When human verification is required, the workflow should stop for manual intervention.
