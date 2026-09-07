from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import json
import re


# -------------------------------------------------------------------
# Resume tailoring prompt
# -------------------------------------------------------------------

GENERIC_RESUME_TEX = r"""
\documentclass[11pt,letterpaper]{article}

\usepackage[margin=0.55in]{geometry}
\usepackage{titlesec}
\usepackage{enumitem}
\usepackage{hyperref}
\usepackage{xcolor}
\usepackage{mathptmx}
\usepackage{tabularx}

\pagestyle{empty}
\setlength{\parindent}{0pt}
\linespread{1.0}

\hypersetup{
    colorlinks=true,
    urlcolor=blue
}

\titleformat{\section}{\large\bfseries}{}{0em}{}
\titlespacing{\section}{0pt}{6pt}{3pt}

\newcommand{\sectionrule}{\vspace{-6pt}\hrule\vspace{3pt}}

\newcommand{\entryHeader}[4]{
    \textbf{#1} \hfill \textit{#2}\\
    \textbf{#3} \hfill \textit{#4}
}

\newenvironment{entryList}{
    \begin{itemize}
        [leftmargin=1.3em,itemsep=1.5pt,topsep=2pt,parsep=0pt]
}{
    \end{itemize}
    \vspace{2pt}
}

\begin{document}

\begin{center}
{\LARGE \bfseries Priya Yadav}\\[2pt]
{\normalsize Backend Software Engineer}\\[2pt]
\small +91 9350046768 ~$\bullet$~
\href{mailto:priya2668135@gmail.com}{priya2668135@gmail.com} ~$\bullet$~
\href{https://www.linkedin.com/in/priya-yadav-14a44419/}{LinkedIn} ~$\bullet$~
\href{https://github.com/precise-charmerr}{GitHub}
\end{center}

\vspace{1pt}

\section*{PROFILE}
\sectionrule

Backend Software Engineer with 3 years of experience designing, implementing, and optimizing high-performance microservices, REST APIs, and distributed data pipelines using Java, Spring Boot, Python, and SQL. Strong background in distributed systems, resilient backend services, and API development, with hands-on experience in Redis, CI/CD, and Agile engineering.

\section*{WORK EXPERIENCE}
\sectionrule

\entryHeader{Software Developer}{Jun 2023 -- Jul 2026}{Deutsche Bank}{Pune, Maharashtra}

\begin{entryList}

\item Designed, tuned, and deployed scalable RESTful APIs with FastAPI/Flask and Java Spring Boot to integrate complex data ingestion and Parquet-based analytics across internal services.

\item Architected and implemented a high-availability Redis distributed caching layer to decouple core authorization services, ensuring service reliability and up to 3 days of outage-free operational continuity.

\item Optimized data lifecycle management and PySpark ETL pipelines handling 20M+ large data records, cutting pipeline execution runtime and improving database read/write performance.

\item Collaborated closely with cross-functional Agile engineering teams, participating in sprint planning, code reviews, and CI/CD automation to uphold robust software quality and security standards.

\item Contributed to system observability and on-call rotation readiness, proactively identifying and resolving backend service bottlenecks to maintain high system availability.

\end{entryList}

\section*{PROJECTS}
\sectionrule

\entryHeader{Document Q\&A Agent (AI-Powered RAG Service)}{Present}{Independent Project}{Pune, Maharashtra}

\begin{entryList}

\item Developing an agentic Retrieval-Augmented Generation (RAG) backend utilizing Python, FastAPI, and LangGraph to automate document context evaluation and multi-step reasoning.

\item Advanced structured prompt engineering and system context management to drive accurate, secure, and production-ready LLM outputs.

\item Designed vector search infrastructure (FAISS/Chroma) to efficiently embed, index, and query unstructured data, exposing functionality via clean REST API endpoints.

\end{entryList}

\section*{SKILLS}
\sectionrule

\begin{tabularx}{\textwidth}{@{}p{1.85in}X@{}}

\textbf{Languages} & Java $\bullet$ Python $\bullet$ SQL $\bullet$ JavaScript $\bullet$ HTML/CSS \\

\textbf{Backend \& Cloud Tech} & Spring Boot $\bullet$ Microservices $\bullet$ Spring MVC $\bullet$ RESTful APIs $\bullet$ Redis $\bullet$ PySpark $\bullet$ FastAPI $\bullet$ Flask $\bullet$ Spring Security \\

\textbf{AI \& Prompt Engineering} & Prompt Engineering $\bullet$ RAG Architecture $\bullet$ LangGraph $\bullet$ Vector Search (FAISS/Chroma) $\bullet$ LLM Integrations \\

\textbf{Tools, DevOps \& Testing} & Git $\bullet$ CI/CD Pipelines $\bullet$ Maven $\bullet$ Postman $\bullet$ MS SQL Server $\bullet$ Elasticsearch $\bullet$ BDD Practices \\

\textbf{Methodologies \& Concepts} & Code Review \& Security $\bullet$ Agile / Scrum $\bullet$ Distributed Systems $\bullet$ Multithreading $\bullet$ System Design $\bullet$ Object-Oriented Programming (OOP) \\

\end{tabularx}

\section*{EDUCATION \& ACHIEVEMENTS}
\sectionrule

\entryHeader{BE in Computer Engineering}{Jul 2019 -- Jun 2023}{Army Institute of Technology}{Pune, Maharashtra}

\begin{entryList}

\item GPA: 8.74/10 | Relevant Coursework: Data Structures \& Algorithms, DBMS, OOP, System Design, Software Engineering

\item \textbf{Honors \& Mentorship:} Google WE (Women Engineers) Program Scholar $\bullet$ Microsoft Engage Mentee

\end{entryList}

\end{document}
"""


RESUME_PROMPT = r"""
You are an expert ATS resume writer.

Your task is to create a tailored one-page LaTeX resume for the candidate
based ONLY on the candidate's original resume and the provided Job
Description.

The candidate's original resume is the absolute source of truth.

======================================================================
ORIGINAL RESUME
======================================================================

{{GENERIC_RESUME_TEX}}

======================================================================
END ORIGINAL RESUME
======================================================================

Analyze the Job Description carefully and create a tailored version of
the original resume.

The goal is to maximize the truthful match between the candidate's
existing background and this specific Job Description.

======================================================================
1. IDENTIFY THE STRONGEST MATCH
======================================================================

Before generating the resume, internally analyze:

- Which technical skills in the resume directly match the JD?
- Which work-experience responsibilities are most relevant?
- Which projects are most relevant?
- Which technologies mentioned in the JD are already present in the resume?
- Which JD requirements are not present in the resume?
- Which keywords are important for ATS matching?
- What type of role is this primarily:
  Java/Spring backend,
  Python backend,
  distributed systems,
  data engineering,
  AI/LLM,
  or another backend-focused role?

Use this analysis to decide what should receive the most emphasis.

Do not output this analysis.

======================================================================
2. PRIORITIZE RELEVANT SKILLS
======================================================================

Do NOT give equal emphasis to every skill.

Prioritize the skills that are most relevant to this specific JD.

For example:

If the JD emphasizes Java, Spring Boot, Spring MVC, Spring Security,
microservices, REST APIs, SQL, and backend development, prioritize the
candidate's Java/Spring Boot experience.

If the JD emphasizes Python, FastAPI, Flask, PySpark, data processing,
or data pipelines, prioritize the candidate's Python/data experience.

If the JD emphasizes distributed systems, high availability, caching,
scalability, or large-scale data, emphasize the candidate's Redis
distributed caching, microservices, distributed systems, API, and
20M+ record data-processing experience.

If the JD emphasizes AI, LLMs, RAG, vector databases, LangGraph, or
prompt engineering, emphasize the Document Q&A Agent project and the
AI-related skills actually present in the resume.

These are examples only.

Always make prioritization based on the actual JD.

======================================================================
3. TAILOR THE PROFILE
======================================================================

Rewrite the PROFILE section so that it presents the candidate as the
strongest truthful candidate for this specific role.

Emphasize experience and technologies that are most relevant to the JD.

Use JD terminology when it accurately corresponds to experience already
supported by the original resume.

Do not add technologies, experience, metrics, or claims that are not
supported by the original resume.

Keep approximately 3 years of professional experience.

======================================================================
4. TAILOR WORK EXPERIENCE
======================================================================

Prioritize and rewrite the existing work-experience bullet points based
on relevance to the JD.

You may:

- Reorder existing bullets.
- Rewrite bullets for clarity and stronger technical relevance.
- Emphasize technologies important to the JD when those technologies are
  already present in the candidate's experience.
- Use terminology from the JD when it accurately describes work actually
  performed.
- Make relevant technical details more prominent.
- Make wording slightly more concise when necessary to fit one page.

IMPORTANT:

The original resume contains five substantive work-experience bullets.

Preserve all five work-experience bullets.

Do not remove work-experience bullets merely because they are less
relevant to the JD.

You may change their order and wording, but preserve their factual
meaning.

Do NOT:

- Invent responsibilities.
- Invent technologies.
- Invent metrics.
- Invent achievements.
- Change employment dates.
- Change the employer.
- Change the job title.
- Change the location.
- Claim production experience with a technology merely because it appears
  in the JD.

For example, if the JD requires Spring Boot and the resume already
contains Spring Boot experience, make that experience prominent.

If the JD requires Kafka but Kafka does not appear in the resume,
do not add Kafka.

======================================================================
5. TAILOR PROJECTS
======================================================================

Keep the existing Document Q&A Agent project.

Do NOT remove the project simply because it is less relevant to the JD.

If it is relevant, emphasize the parts that best match the role, such as:

- Python
- FastAPI
- LangGraph
- RAG
- vector search
- FAISS/Chroma
- REST APIs
- prompt engineering
- LLM integrations

If the JD is primarily a traditional Java/Spring backend role and the
AI project is less relevant, do not allow the AI project to overshadow
the professional Java/backend experience.

You may reorder or rewrite the project bullets for relevance, but
preserve all three substantive project bullets and their factual meaning.

======================================================================
6. TAILOR SKILLS
======================================================================

Reorganize the SKILLS section so that the most relevant skills appear
first within each category.

Only include skills that are already supported by the original resume.

Use the JD's terminology when it accurately corresponds to an existing
skill.

Do not add missing JD technologies simply to improve ATS matching.

Do not keyword-stuff the resume.

Keep all existing skill categories and substantive skills.

======================================================================
7. TRUTHFULNESS
======================================================================

This is extremely important.

The original resume is the ONLY source of truth.

Never invent or imply that the candidate has experience with a technology
simply because it appears in the JD.

For example, if the JD says:

"Experience with Kafka, Kubernetes, and AWS required"

but the original resume does not contain Kafka, Kubernetes, or AWS,
do NOT add them.

Instead, maximize the match using the experience that actually exists,
such as Java, Spring Boot, microservices, REST APIs, Redis, SQL,
distributed systems, CI/CD, etc., if relevant.

Do not fabricate missing experience.

Do not convert JD requirements into candidate skills.

Do not infer experience merely because two technologies are commonly
used together.

======================================================================
8. ATS OPTIMIZATION
======================================================================

Optimize the resume for ATS by:

- Using relevant keywords naturally.
- Prioritizing exact terminology from the JD where truthful.
- Making important technical skills easy for ATS systems to identify.
- Avoiding keyword stuffing.
- Keeping standard resume section names.
- Maintaining a clean, machine-readable structure.
- Preserving the existing LaTeX resume template.

Do not sacrifice truthfulness for ATS matching.

======================================================================
9. RESUME STRUCTURE
======================================================================

Keep the existing resume structure and formatting.

Do not unnecessarily redesign the resume.

Maintain:

- Header
- Profile
- Work Experience
- Projects
- Skills
- Education & Achievements

You may modify the content and ordering within these sections to improve
relevance.

Do not remove an entire section.

Do not introduce:

- graphics
- icons
- photographs
- sidebars
- decorative elements
- unnecessary columns
- ATS-unfriendly layouts

Use the provided generic LaTeX resume as the structural and formatting
baseline.

======================================================================
10. ONE-PAGE AND PAGE-FILL REQUIREMENT
======================================================================

The final resume MUST fit on exactly ONE letter-size page.

This is mandatory.

The resume should use the available page height efficiently and should
have approximately the same visual density as the original generic
resume.

Do NOT create a sparse resume with a large unused blank area at the
bottom.

Do NOT solve the one-page requirement by deleting substantial factual
content.

The complete resume should retain:

- Profile
- all five Work Experience bullets
- Projects
- all three Project bullets
- Skills
- Education & Achievements

If the content is too long to fit on one page, first optimize:

- wording
- unnecessary whitespace
- section spacing
- bullet spacing
- line spacing
- small formatting details

You may make small ATS-friendly formatting adjustments when necessary.

Reasonable adjustments include:

- slightly reducing section spacing
- slightly reducing bullet spacing
- slightly reducing line spacing
- slightly adjusting margins within reasonable ATS-friendly limits
- making bullets more concise while preserving factual meaning

Do NOT:

- create a second page
- leave a large unnecessary blank area at the bottom
- use an unreasonably small font
- remove Work Experience
- remove Projects
- remove Skills
- remove Education & Achievements
- remove multiple substantive work bullets
- invent content merely to fill space

The final result must be:

ONE PAGE
+
WELL FILLED
+
READABLE
+
ATS FRIENDLY
+
TRUTHFUL

======================================================================
11. DO NOT EXAGGERATE SENIORITY
======================================================================

The resume represents approximately 3 years of professional experience.

Do not make the candidate appear to have more experience than they
actually have.

Do not use "Senior", "Lead", "Architect", "Principal", or similar
seniority claims unless they are actually supported by the resume and
appropriate for the role.

Do not change the candidate's actual job title.

======================================================================
12. LATEX OUTPUT
======================================================================

Return ONLY the complete LaTeX source code for the tailored resume.

Do not provide:

- an explanation
- match analysis
- a summary
- recommendations
- Markdown code fences
- commentary before or after the resume

The output must be a complete compilable LaTeX document containing:

- \documentclass
- all required packages
- \begin{document}
- the complete tailored resume
- \end{document}

Correctly escape all LaTeX special characters.

Preserve working hyperlinks.

Preserve the formatting and overall visual structure of the original
resume.

======================================================================
13. FINAL INTERNAL VALIDATION
======================================================================

Before returning the answer, internally verify ALL of the following:

1. The output starts with \documentclass.
2. The output ends with \end{document}.
3. The document is complete LaTeX.
4. All required packages are included.
5. The document is compilable LaTeX.
6. The resume is exactly ONE letter-size page.
7. There is no second page.
8. There is no large unused blank area at the bottom.
9. The page is visually well-filled.
10. The Work Experience section is retained.
11. All five original work-experience bullets are retained.
12. The Projects section is retained.
13. All three original project bullets are retained.
14. The Skills section is retained.
15. Education & Achievements is retained.
16. The candidate still represents approximately 3 years of experience.
17. No unsupported technology was added.
18. No responsibility was invented.
19. No metric was invented.
20. No achievement was invented.
21. No employer, job title, date, or location was changed.
22. JD keywords are used only where truthful.
23. The resume remains ATS-friendly.
24. The output contains ONLY the LaTeX source.

======================================================================
JOB DESCRIPTION
======================================================================

{{JOB_DESCRIPTION}}
"""


# -------------------------------------------------------------------
# Experience filtering
# -------------------------------------------------------------------

def experience_ranges(text):
    """
    Extract explicit experience requirements from text.

    Examples:
        "3-5 years"      -> [(3, 5)]
        "2 to 4 years"   -> [(2, 4)]
        "5+ years"       -> [(5, None)]
        "3 years"        -> [(3, 3)]
    """
    if not text:
        return []

    t = text.lower().replace("–", "-").replace("—", "-")

    ranges = []

    # Examples: 3-5 years, 3 to 5 years
    for match in re.finditer(
        r"\b(\d+)\s*(?:-|to)\s*(\d+)\s*(?:years?|yrs?)\b",
        t,
    ):
        minimum = int(match.group(1))
        maximum = int(match.group(2))
        ranges.append((minimum, maximum))

    # Examples: 5+ years, 5 or more years
    for match in re.finditer(
        r"\b(\d+)\s*(?:\+|or\s+more)\s*(?:years?|yrs?)\b",
        t,
    ):
        minimum = int(match.group(1))
        ranges.append((minimum, None))

    # Examples: 3 years, 3 yrs
    for match in re.finditer(
        r"\b(\d+)\s*(?:years?|yrs?)\b",
        t,
    ):
        value = int(match.group(1))

        # Don't duplicate numbers that were already captured as part
        # of a range such as "3-5 years".
        if not any(
            minimum == value
            and (
                maximum == value
                or maximum is None
            )
            for minimum, maximum in ranges
        ):
            ranges.append((value, value))

    return ranges


def eligible(job, max_exp):
    """
    A job is eligible only if every explicit experience requirement
    found in its Experience field and Description is within max_exp.

    Jobs with no explicit experience requirement are allowed.
    """

    experience_text = job.get("experience", "") or ""
    description_text = job.get("description", "") or ""

    combined_text = f"{experience_text} {description_text}"

    ranges = experience_ranges(combined_text)

    if not ranges:
        return True

    for minimum, maximum in ranges:

        # "6+ years" -> reject
        if minimum > max_exp:
            return False

        # "3-7 years" -> reject because maximum exceeds limit
        if maximum is not None and maximum > max_exp:
            return False

    return True


# -------------------------------------------------------------------
# Deduplication
# -------------------------------------------------------------------

def dedupe(jobs):
    seen = set()
    out = []

    for j in jobs:
        key = (j.get("url") or "").strip().lower()

        if not key:
            key = (
                (j.get("title") or "")
                + "|"
                + (j.get("company") or "")
            ).lower()

        if key not in seen:
            seen.add(key)
            out.append(j)

    return out


# -------------------------------------------------------------------
# Ranking
# -------------------------------------------------------------------

def rank(job):
    text = " ".join(
        str(job.get(k, ""))
        for k in ("title", "description", "skills")
    ).lower()

    score = 0

    for s in [
        "java",
        "spring boot",
        "spring",
        "microservices",
        "rest api",
        "sql",
        "mysql",
        "postgresql",
        "kafka",
        "aws",
        "docker",
        "kubernetes",
    ]:
        if s in text:
            score += 2

    title = (job.get("title") or "").lower()

    for s in [
        "java backend developer",
        "java developer",
        "backend developer",
        "software engineer",
    ]:
        if s in title:
            score += 4

    for s in [
        "senior manager",
        "architect",
        "principal",
        "director",
        "android",
        "frontend",
        "ui developer",
        "ios",
    ]:
        if s in title:
            score -= 10

    return score


# -------------------------------------------------------------------
# Resume prompt generation
# -------------------------------------------------------------------

def build_resume_prompt(job):
    """
    Creates the complete resume-generation prompt that will be placed
    into the Resume Prompt column for this specific job.
    """

    description = job.get("description", "") or ""

    return (
        RESUME_PROMPT
        .replace("{{GENERIC_RESUME_TEX}}", GENERIC_RESUME_TEX)
        .replace("{{JOB_DESCRIPTION}}", description)
    )


# -------------------------------------------------------------------
# Excel output
# -------------------------------------------------------------------

def write_excel(jobs, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"

    headers = [
        "Rank",
        "Source",
        "Title",
        "Company",
        "Location",
        "Experience",
        "Skills",
        "Description",
        "URL",
        "Resume Prompt",
    ]

    ws.append(headers)

    for i, job in enumerate(jobs, 1):
        ws.append([
            i,
            job.get("source", ""),
            job.get("title", ""),
            job.get("company", ""),
            job.get("location", ""),
            job.get("experience", ""),
            job.get("skills", ""),
            job.get("description", ""),
            "",
            build_resume_prompt(job),
        ])

        # Make the job URL clickable
        url_cell = ws.cell(row=ws.max_row, column=9)
        url = job.get("url", "")

        if url:
            url_cell.value = "Open Job"
            url_cell.hyperlink = url
            url_cell.style = "Hyperlink"

    # Adjust column widths
    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)

        max_len = max(
            [
                len(str(ws.cell(r, col).value or ""))
                for r in range(1, ws.max_row + 1)
            ]
            or [10]
        )

        ws.column_dimensions[letter].width = min(
            max(max_len + 2, 12),
            55,
        )

    # Keep the spreadsheet easy to navigate
    ws.freeze_panes = "A2"

    # Wrap long JD/prompt cells
    for row in ws.iter_rows():
        for cell in row:
            if cell.column in (8, 10):
                cell.alignment = cell.alignment.copy(
                    wrap_text=True,
                    vertical="top",
                )

    wb.save(path)


# -------------------------------------------------------------------
# Main pipeline
# -------------------------------------------------------------------

def combine_and_write(jobs, config):
    print("\nCombining results...")

    jobs = dedupe(jobs)

    jobs = [
        j
        for j in jobs
        if eligible(
            j,
            config["max_experience_requirement"],
        )
    ]

    jobs.sort(key=rank, reverse=True)

    out_dir = Path("jobs")
    out_dir.mkdir(exist_ok=True)

    xlsx = out_dir / "job-tracker.xlsx"
    js = out_dir / "jobs.json"

    write_excel(jobs, xlsx)

    js.write_text(
        json.dumps(
            jobs,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print(f"Final jobs: {len(jobs)}")
    print(f"Excel: {xlsx}")
    print(f"JSON: {js}")