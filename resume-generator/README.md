# Resume Generator

Generates tailored one-page PDF resumes from the `Tailored Resume .tex` column in `job-tracker.xlsx`.

## Structure

```text
resume-generator/
├── input/
│   └── job-tracker.xlsx
├── resumes/
├── pdfs/
├── logs/
├── generate_resumes.py
├── requirements.txt
└── README.md
```

## Setup

```bash
cd ~/job_automation/resume-generator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Make sure LaTeX is available:

```bash
which pdflatex
pdflatex --version
```

Put `job-tracker.xlsx` in `input/`.

The workbook must contain `Rank` and `Tailored Resume .tex`.

## Run

```bash
python generate_resumes.py
```

Output names:

```text
Priya_Resume_3yrs(1).tex
Priya_Resume_3yrs(1).pdf
Priya_Resume_3yrs(37).tex
Priya_Resume_3yrs(37).pdf
```

The script does not open a browser, access LinkedIn/Naukri, submit applications, or modify the Excel workbook.
