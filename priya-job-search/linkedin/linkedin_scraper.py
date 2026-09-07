import re
import time
import requests
from bs4 import BeautifulSoup


SEARCH_URL = (
    "https://www.linkedin.com/jobs-guest/jobs/api/"
    "seeMoreJobPostings/search"
)

DETAIL_URL = (
    "https://www.linkedin.com/jobs-guest/jobs/api/"
    "jobPosting/{}"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    )
}


def clean_text(element):
    if not element:
        return ""

    return " ".join(
        element.get_text(" ", strip=True).split()
    )


def extract_job_id(card):
    """
    Extract LinkedIn job ID from:
    data-entity-urn="urn:li:jobPosting:123456789"
    """

    urn = card.get("data-entity-urn", "")

    match = re.search(
        r"jobPosting:(\d+)",
        urn
    )

    if match:
        return match.group(1)

    # Fallback: sometimes the attribute is on a child element
    element = card.select_one(
        "[data-entity-urn*='jobPosting:']"
    )

    if element:
        urn = element.get(
            "data-entity-urn",
            ""
        )

        match = re.search(
            r"jobPosting:(\d+)",
            urn
        )

        if match:
            return match.group(1)

    # Final fallback: extract ID from job URL
    link = card.select_one(
        "a.base-card__full-link"
    )

    if link:
        href = link.get("href", "")

        match = re.search(
            r"/view/(\d+)",
            href
        )

        if match:
            return match.group(1)

    return ""


def fetch_description(session, job_id):
    """
    Fetch the individual LinkedIn job page and extract
    the full job description.

    Includes retry/backoff for 429 and 5xx.
    """

    if not job_id:
        return ""

    url = DETAIL_URL.format(job_id)

    max_attempts = 3

    for attempt in range(max_attempts):

        try:
            response = session.get(
                url,
                timeout=30
            )

            if response.status_code == 200:

                soup = BeautifulSoup(
                    response.text,
                    "html.parser"
                )

                selectors = [
                    ".show-more-less-html__markup",
                    ".description__text"
                ]

                for selector in selectors:

                    element = soup.select_one(
                        selector
                    )

                    description = clean_text(
                        element
                    )

                    if len(description) >= 50:
                        return description

                return ""

            if response.status_code in (429, 500, 502, 503, 504):

                wait_seconds = 2 ** attempt

                print(
                    f"      JD HTTP "
                    f"{response.status_code}, "
                    f"retrying in {wait_seconds}s..."
                )

                time.sleep(wait_seconds)
                continue

            print(
                f"      JD HTTP "
                f"{response.status_code}"
            )

            return ""

        except requests.RequestException as e:

            if attempt == max_attempts - 1:
                print(
                    f"      JD request failed: {e}"
                )
                return ""

            wait_seconds = 2 ** attempt

            print(
                f"      JD request error, "
                f"retrying in {wait_seconds}s..."
            )

            time.sleep(wait_seconds)

    return ""


def run_linkedin(config):

    jobs = []

    queries = config["queries"]
    locations = config["locations"]

    pages_per_search = config.get(
        "pages_per_search",
        1
    )

    results_per_search = config.get(
        "results_per_search",
        25
    )

    # Last 24 hours
    posted_within = config.get(
        "linkedin_posted_within",
        "r86400"
    )

    session = requests.Session()

    session.headers.update(
        HEADERS
    )

    for query in queries:

        for location in locations:

            print(
                f"  LinkedIn: "
                f"{query} | {location}"
            )

            captured_for_search = 0

            for page_number in range(
                pages_per_search
            ):

                if (
                    captured_for_search
                    >= results_per_search
                ):
                    break

                start = page_number * 10

                params = {
                    "keywords": query,
                    "location": location,
                    "f_TPR": posted_within,
                    "start": start
                }

                try:

                    response = session.get(
                        SEARCH_URL,
                        params=params,
                        timeout=30
                    )

                    if response.status_code == 429:

                        print(
                            "    LinkedIn HTTP 429 "
                            "(rate limited)"
                        )

                        wait_seconds = 10

                        print(
                            f"    Waiting "
                            f"{wait_seconds}s..."
                        )

                        time.sleep(
                            wait_seconds
                        )

                        continue

                    if response.status_code != 200:

                        print(
                            f"    LinkedIn HTTP "
                            f"{response.status_code}"
                        )

                        continue

                    soup = BeautifulSoup(
                        response.text,
                        "html.parser"
                    )

                    cards = soup.select(
                        "li"
                    )

                    for card in cards:

                        if (
                            captured_for_search
                            >= results_per_search
                        ):
                            break

                        title_el = card.select_one(
                            ".base-search-card__title"
                        )

                        company_el = card.select_one(
                            ".base-search-card__subtitle"
                        )

                        location_el = card.select_one(
                            ".job-search-card__location"
                        )

                        link_el = card.select_one(
                            "a.base-card__full-link"
                        )

                        title = clean_text(
                            title_el
                        )

                        company = clean_text(
                            company_el
                        )

                        job_location = (
                            clean_text(
                                location_el
                            )
                            or location
                        )

                        link = ""

                        if link_el:

                            link = (
                                link_el
                                .get("href", "")
                                .strip()
                            )

                        if not title:
                            continue

                        # ---------------------------------
                        # Get LinkedIn job ID
                        # ---------------------------------

                        job_id = extract_job_id(
                            card
                        )

                        # ---------------------------------
                        # Fetch JD
                        # ---------------------------------

                        description = ""

                        if job_id:

                            print(
                                f"    Fetching JD: "
                                f"{title}"
                            )

                            description = (
                                fetch_description(
                                    session,
                                    job_id
                                )
                            )

                            if description:

                                print(
                                    f"      JD captured: "
                                    f"{len(description)} "
                                    f"chars"
                                )

                            else:

                                print(
                                    "      JD not available"
                                )

                            # Small delay between
                            # detail requests
                            time.sleep(1)

                        else:

                            print(
                                "      Job ID not found"
                            )

                        jobs.append({
                            "source": "LinkedIn",
                            "title": title,
                            "company": company,
                            "location": job_location,
                            "experience": "",
                            "skills": "",
                            "description": description,
                            "url": link,
                        })

                        captured_for_search += 1

                except requests.RequestException as e:

                    print(
                        f"    skipped page "
                        f"{page_number + 1}: {e}"
                    )

            print(
                f"    captured: "
                f"{captured_for_search}"
            )

    session.close()

    return jobs