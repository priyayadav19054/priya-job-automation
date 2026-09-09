import re
import time
import requests
from bs4 import BeautifulSoup


# -------------------------------------------------------------------
# LinkedIn endpoints
# -------------------------------------------------------------------

SEARCH_URL = (
    "https://www.linkedin.com/jobs-guest/jobs/api/"
    "seeMoreJobPostings/search"
)

DETAIL_URL = (
    "https://www.linkedin.com/jobs-guest/jobs/api/"
    "jobPosting/{}"
)


# -------------------------------------------------------------------
# HTTP configuration
# -------------------------------------------------------------------

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    )
}


# -------------------------------------------------------------------
# Timing configuration
# -------------------------------------------------------------------

# Delay between successful JD requests.
JD_DELAY_SECONDS = 1.5

# When LinkedIn returns 429, use progressively longer waits.
JD_429_BACKOFF = [30, 60, 120]

# Search-page 429 backoff.
SEARCH_429_BACKOFF = [30, 60, 120]

# Maximum number of consecutive search-page 429s before giving up
# on the current query/location.
MAX_CONSECUTIVE_SEARCH_429 = 3

# Maximum number of consecutive JD 429s before we stop trying to
# fetch more JDs for the current run.
MAX_CONSECUTIVE_JD_429 = 3


# -------------------------------------------------------------------
# Text helpers
# -------------------------------------------------------------------

def clean_text(element):
    if not element:
        return ""

    return " ".join(
        element.get_text(
            " ",
            strip=True
        ).split()
    )


# -------------------------------------------------------------------
# Job ID extraction
# -------------------------------------------------------------------

def extract_job_id(card):
    """
    Extract LinkedIn job ID from:

        data-entity-urn="urn:li:jobPosting:123456789"

    Falls back to a child element or job URL.
    """

    urn = card.get(
        "data-entity-urn",
        ""
    )

    match = re.search(
        r"jobPosting:(\d+)",
        urn
    )

    if match:
        return match.group(1)

    # ---------------------------------------------------------------
    # Fallback: sometimes the attribute is on a child element.
    # ---------------------------------------------------------------

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

    # ---------------------------------------------------------------
    # Final fallback: extract ID from job URL.
    # ---------------------------------------------------------------

    link = card.select_one(
        "a.base-card__full-link"
    )

    if link:

        href = link.get(
            "href",
            ""
        )

        match = re.search(
            r"/view/(\d+)",
            href
        )

        if match:
            return match.group(1)

    return ""


# -------------------------------------------------------------------
# JD fetching
# -------------------------------------------------------------------

def fetch_description(session, job_id):
    """
    Fetch the individual LinkedIn job page and extract
    the full job description.

    429 handling uses long exponential backoff rather than
    repeatedly retrying after 1-2-4 seconds.

    Returns:

        description -> successful JD
        ""          -> unavailable / rate limited / failed
    """

    if not job_id:
        return ""

    url = DETAIL_URL.format(
        job_id
    )

    max_attempts = len(
        JD_429_BACKOFF
    )

    for attempt in range(
        max_attempts
    ):

        try:

            response = session.get(
                url,
                timeout=30
            )

            # -------------------------------------------------------
            # Successful response
            # -------------------------------------------------------

            if response.status_code == 200:

                soup = BeautifulSoup(
                    response.text,
                    "html.parser"
                )

                selectors = [
                    ".show-more-less-html__markup",
                    ".description__text",
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

            # -------------------------------------------------------
            # Rate limited
            # -------------------------------------------------------

            if response.status_code == 429:

                wait_seconds = JD_429_BACKOFF[
                    min(
                        attempt,
                        len(JD_429_BACKOFF) - 1
                    )
                ]

                print(
                    f"      JD HTTP 429, "
                    f"waiting {wait_seconds}s "
                    f"before retry..."
                )

                time.sleep(
                    wait_seconds
                )

                continue

            # -------------------------------------------------------
            # Temporary server errors
            # -------------------------------------------------------

            if response.status_code in (
                500,
                502,
                503,
                504,
            ):

                wait_seconds = 10 * (
                    attempt + 1
                )

                print(
                    f"      JD HTTP "
                    f"{response.status_code}, "
                    f"waiting {wait_seconds}s "
                    f"before retry..."
                )

                time.sleep(
                    wait_seconds
                )

                continue

            # -------------------------------------------------------
            # Other HTTP errors
            # -------------------------------------------------------

            print(
                f"      JD HTTP "
                f"{response.status_code}"
            )

            return ""

        except requests.RequestException as e:

            if attempt == max_attempts - 1:

                print(
                    f"      JD request failed: "
                    f"{e}"
                )

                return ""

            wait_seconds = 10 * (
                attempt + 1
            )

            print(
                f"      JD request error, "
                f"waiting {wait_seconds}s "
                f"before retry..."
            )

            time.sleep(
                wait_seconds
            )

    return ""


# -------------------------------------------------------------------
# Search request
# -------------------------------------------------------------------

def fetch_search_page(
    session,
    query,
    location,
    posted_within,
    start,
):
    """
    Fetch one LinkedIn search-results page.

    Returns:

        soup
            successful response

        None
            failed/rate-limited response
    """

    params = {
        "keywords": query,
        "location": location,
        "f_TPR": posted_within,
        "start": start,
    }

    try:

        response = session.get(
            SEARCH_URL,
            params=params,
            timeout=30
        )

        if response.status_code == 200:

            return BeautifulSoup(
                response.text,
                "html.parser"
            )

        print(
            f"    LinkedIn HTTP "
            f"{response.status_code}"
        )

        return None

    except requests.RequestException as e:

        print(
            f"    LinkedIn search request "
            f"failed: {e}"
        )

        return None


# -------------------------------------------------------------------
# Main LinkedIn scraper
# -------------------------------------------------------------------

def run_linkedin(config):

    jobs = []

    queries = config[
        "queries"
    ]

    locations = config[
        "locations"
    ]

    pages_per_search = config.get(
        "pages_per_search",
        1
    )

    results_per_search = config.get(
        "results_per_search",
        25
    )

    # Last 24 hours.
    posted_within = config.get(
        "linkedin_posted_within",
        "r86400"
    )

    session = requests.Session()

    session.headers.update(
        HEADERS
    )

    # Track JD rate limiting globally.
    consecutive_jd_429 = 0

    # ---------------------------------------------------------------
    # Query / location loop
    # ---------------------------------------------------------------

    stop_linkedin = False

    for query in queries:

        if stop_linkedin:
            break

        for location in locations:

            if stop_linkedin:
                break

            print(
                f"  LinkedIn: "
                f"{query} | {location}"
            )

            captured_for_search = 0

            # Track search-page 429s independently for this
            # query/location.
            consecutive_search_429 = 0

            # -------------------------------------------------------
            # Pages
            # -------------------------------------------------------

            for page_number in range(
                pages_per_search
            ):

                if (
                    captured_for_search
                    >= results_per_search
                ):
                    break

                start = (
                    page_number * 10
                )

                # ---------------------------------------------------
                # Search request with retry/backoff
                # ---------------------------------------------------

                soup = None

                for search_attempt in range(
                    len(SEARCH_429_BACKOFF)
                ):

                    params = {
                        "keywords": query,
                        "location": location,
                        "f_TPR": posted_within,
                        "start": start,
                    }

                    try:

                        response = session.get(
                            SEARCH_URL,
                            params=params,
                            timeout=30
                        )

                        # -------------------------------------------
                        # Success
                        # -------------------------------------------

                        if response.status_code == 200:

                            soup = BeautifulSoup(
                                response.text,
                                "html.parser"
                            )

                            consecutive_search_429 = 0

                            break

                        # -------------------------------------------
                        # Rate limited
                        # -------------------------------------------

                        if response.status_code == 429:

                            consecutive_search_429 += 1

                            wait_seconds = (
                                SEARCH_429_BACKOFF[
                                    search_attempt
                                ]
                            )

                            print(
                                "    LinkedIn HTTP 429 "
                                "(rate limited)"
                            )

                            print(
                                f"    Waiting "
                                f"{wait_seconds}s "
                                f"before retry..."
                            )

                            time.sleep(
                                wait_seconds
                            )

                            continue

                        # -------------------------------------------
                        # Other HTTP error
                        # -------------------------------------------

                        print(
                            f"    LinkedIn HTTP "
                            f"{response.status_code}"
                        )

                        break

                    except requests.RequestException as e:

                        print(
                            f"    LinkedIn search "
                            f"request failed: {e}"
                        )

                        break

                # ---------------------------------------------------
                # If search page still failed, decide whether to
                # continue or stop LinkedIn.
                # ---------------------------------------------------

                if soup is None:

                    if (
                        consecutive_search_429
                        >= MAX_CONSECUTIVE_SEARCH_429
                    ):

                        print(
                            "\n    LinkedIn is "
                            "currently rate limiting "
                            "the scraper."
                        )

                        print(
                            "    Stopping LinkedIn "
                            "for this run."
                        )

                        stop_linkedin = True

                    break

                # ---------------------------------------------------
                # Extract job cards
                # ---------------------------------------------------

                cards = soup.select(
                    "li"
                )

                # ---------------------------------------------------
                # Process cards
                # ---------------------------------------------------

                for card in cards:

                    if (
                        captured_for_search
                        >= results_per_search
                    ):
                        break

                    # -----------------------------------------------
                    # Basic fields
                    # -----------------------------------------------

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
                            .get(
                                "href",
                                ""
                            )
                            .strip()
                        )

                    if not title:
                        continue

                    # -----------------------------------------------
                    # Get LinkedIn job ID
                    # -----------------------------------------------

                    job_id = extract_job_id(
                        card
                    )

                    # -----------------------------------------------
                    # Fetch JD
                    # -----------------------------------------------

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

                            # Successful request.
                            consecutive_jd_429 = 0

                        else:

                            print(
                                "      JD not available"
                            )

                        # Small delay between detail
                        # requests.
                        time.sleep(
                            JD_DELAY_SECONDS
                        )

                    else:

                        print(
                            "      Job ID not found"
                        )

                    # -----------------------------------------------
                    # Store job
                    # -----------------------------------------------

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

                    # ------------------------------------------------
                    # If we repeatedly hit 429 while fetching JDs,
                    # stop fetching more LinkedIn jobs for this run.
                    #
                    # IMPORTANT:
                    # Already captured jobs remain in `jobs`.
                    # ------------------------------------------------

                    if (
                        consecutive_jd_429
                        >= MAX_CONSECUTIVE_JD_429
                    ):

                        print(
                            "\n    LinkedIn JD endpoint "
                            "is rate limiting heavily."
                        )

                        print(
                            "    Stopping LinkedIn "
                            "for this run."
                        )

                        stop_linkedin = True

                        break

                if stop_linkedin:
                    break

            print(
                f"    captured: "
                f"{captured_for_search}"
            )

    session.close()

    return jobs