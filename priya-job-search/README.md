# Priya Job Search

A local Python pipeline for collecting, filtering, deduplicating, ranking, and tracking jobs from LinkedIn and Naukri.

## Workflow

```text
LinkedIn ─────┐
              ├──→ Job Collection
Naukri ───────┘
                    ↓
              Deduplication
                    ↓
           Experience Filtering
                    ↓
                 Ranking
                    ↓
            Excel + JSON Output
                    ↓
              Resume Prompt
```

## Setup

From the repository root:

```bash
cd priya-job-search
```

Create a virtual environment:

```bash
python -m venv .venv
```

Activate it.

macOS/Linux:

```bash
source .venv/bin/activate
```

Windows:

```powershell
.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Install Playwright Chromium:

```bash
playwright install chromium
```

## Configuration

Configure the project using its existing configuration/environment files.

Typical settings include:

- Job search roles
- Target locations
- Maximum experience requirement
- LinkedIn configuration
- Naukri configuration
- Optional email configuration

Never commit passwords, API keys, cookies, session data, or other secrets.

## Run the Search

The main entry point is:

```bash
python run_all.py
```

The pipeline runs the configured LinkedIn and Naukri collectors and sends their results through the common pipeline.

## Processing

The common pipeline performs:

### Deduplication

Jobs are primarily deduplicated by URL.

When a URL is unavailable, title and company are used as the fallback key.

### Experience Filtering

Explicit experience requirements are extracted from the job's experience information and description.

Jobs exceeding the configured maximum experience requirement are excluded.

Jobs with no explicit experience requirement are allowed.

### Ranking

Jobs are ranked using relevant backend/software-engineering terms, including technologies and role terms such as Java, Spring Boot, Spring, Microservices, REST API, SQL, and backend/software-engineering titles.

## Output

The pipeline creates:

```text
jobs/
├── job-tracker.xlsx
└── jobs.json
```

### Excel Columns

```text
Rank
Source
Title
Company
Location
Experience
Skills
Description
URL
Resume Prompt
```

The URL cell is stored as a clickable Excel hyperlink.

## Resume Prompt

Every job receives a job-specific `Resume Prompt`.

The prompt includes:

- The complete original LaTeX resume
- The job description
- ATS tailoring instructions
- Truthfulness requirements
- One-page requirements
- LaTeX output requirements

The prompt instructs the model to use the original resume as the only source of truth.

A JD technology must not be added to the resume unless the original resume supports it.

## Tailored Resume

For each selected job:

1. Open `job-tracker.xlsx`.
2. Copy the job's `Resume Prompt`.
3. Provide it to the LLM.
4. Save the returned output as `.tex`.
5. Compile the LaTeX.
6. Verify that the final PDF is exactly one page.

The generated resume should preserve the existing structure:

```text
Header
Profile
Work Experience
Projects
Skills
Education & Achievements
```

## Troubleshooting

### Chromium does not start

```bash
playwright install chromium
```

### No jobs are returned

Check:

- Search roles
- Locations
- Experience configuration
- Source configuration
- Authentication/session requirements
- Website changes

### Resume Prompt does not contain the original resume

Confirm that `RESUME_PROMPT` contains:

```text
{{GENERIC_RESUME_TEX}}
```

and that `build_resume_prompt()` replaces it with `GENERIC_RESUME_TEX`.

## Security

Keep credentials and private session information out of Git.

Use environment variables or local configuration for secrets.
