import html
import json
import re
import time

from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


SEARCH_URL = "https://in.indeed.com/jobs"


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean_text(text):
    """
    Normalize whitespace.
    """
    if not text:
        return ""

    return " ".join(str(text).split())


def clean_description(value):
    """
    Convert description HTML/text into clean plain text.
    """

    if value is None:
        return ""

    if isinstance(value, dict):
        # Common possible fields
        for key in [
            "text",
            "html",
            "content",
            "value",
            "description",
            "descriptionText",
            "descriptionHtml",
        ]:
            if key in value:
                result = clean_description(value[key])

                if result:
                    return result

        return ""

    if isinstance(value, list):
        parts = []

        for item in value:
            text = clean_description(item)

            if text:
                parts.append(text)

        return clean_text(" ".join(parts))

    text = str(value)

    # Decode escaped HTML entities
    text = html.unescape(text)

    # Convert common HTML line breaks to spaces
    text = re.sub(
        r"<br\s*/?>",
        "\n",
        text,
        flags=re.IGNORECASE,
    )

    # Remove HTML tags
    text = re.sub(
        r"<[^>]+>",
        " ",
        text,
    )

    return clean_text(text)


# ============================================================
# JOB ID / URL HELPERS
# ============================================================

def extract_job_id(href):
    """
    Extract Indeed job ID from a job URL.

    Example:
        /viewjob?jk=abc123
        /rc/clk?jk=abc123
    """

    if not href:
        return ""

    match = re.search(
        r"[?&]jk=([a-zA-Z0-9]+)",
        href,
    )

    if match:
        return match.group(1)

    return ""


def make_absolute_url(href):
    """
    Convert Indeed relative URLs to absolute URLs.
    """

    if not href:
        return ""

    href = href.strip()

    if href.startswith("http"):
        return href

    if href.startswith("//"):
        return "https:" + href

    if href.startswith("/"):
        return "https://in.indeed.com" + href

    return href


# ============================================================
# CARD HELPERS
# ============================================================

def get_text(card, selectors):
    """
    Try multiple selectors and return the first
    non-empty text.
    """

    for selector in selectors:

        try:

            element = card.locator(
                selector
            ).first

            if element.count() == 0:
                continue

            text = element.inner_text(
                timeout=3000
            )

            text = clean_text(text)

            if text:
                return text

        except Exception:
            continue

    return ""


def get_attribute(card, selectors, attribute):
    """
    Try multiple selectors and return the first
    non-empty attribute.
    """

    for selector in selectors:

        try:

            element = card.locator(
                selector
            ).first

            if element.count() == 0:
                continue

            value = element.get_attribute(
                attribute,
                timeout=3000,
            )

            if value:
                return value.strip()

        except Exception:
            continue

    return ""


def extract_title(card):
    return get_text(
        card,
        [
            "h2.jobTitle span",
            "h3.jobTitle span",
            "h2.jobTitle",
            "h3.jobTitle",
            "a.jcs-JobTitle",
            "[data-testid='job-title']",
        ],
    )


def extract_company(card):
    return get_text(
        card,
        [
            "[data-testid='company-name']",
            ".companyName",
            "[class*='companyName']",
        ],
    )


def extract_location(card):
    return get_text(
        card,
        [
            "[data-testid='text-location']",
            ".companyLocation",
            "[class*='companyLocation']",
        ],
    )


def extract_salary(card):
    """
    Extract salary information when available.
    """

    return get_text(
        card,
        [
            "[data-testid='attribute_snippet_testid salary-snippet-container']",
            ".salary-snippet-container",
            "[class*='salary-snippet']",
        ],
    )


def extract_job_url(card):
    """
    Extract job URL from the job title link.
    """

    href = get_attribute(
        card,
        [
            "a.jcs-JobTitle",
            "h2.jobTitle a",
            "h3.jobTitle a",
            "a[data-jk]",
            "a[href*='jk=']",
            "a[href*='/viewjob']",
        ],
        "href",
    )

    return make_absolute_url(href)


def extract_card_snippet(card):
    """
    Indeed search cards may sometimes expose a short
    snippet. Use it as a fallback only.
    """

    return get_text(
        card,
        [
            ".job-snippet",
            "[class*='job-snippet']",
            "[data-testid='job-snippet']",
        ],
    )


# ============================================================
# JOB CARD DISCOVERY
# ============================================================

def find_job_cards(page):
    """
    Find Indeed job cards using multiple selectors.
    """

    selectors = [
        "div.job_seen_beacon",
        "[data-jk]",
        "[data-testid='slider_item']",
        "[data-testid='job-card']",
        "article",
    ]

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            )

            count = locator.count()

            print(
                f"    Selector {selector}: {count}"
            )

            if count > 0:
                return locator

        except Exception:
            continue

    return None


# ============================================================
# EMBEDDED INDEED DATA
# ============================================================

def get_initial_data(page):
    """
    Read Indeed's window._initialData object directly
    from the browser.

    The search page currently embeds job-detail data
    inside this object.
    """

    try:

        data = page.evaluate(
            """
            () => {
                if (typeof window._initialData === "undefined") {
                    return null;
                }

                return window._initialData;
            }
            """
        )

        return data

    except Exception as e:

        print(
            f"    Could not read window._initialData: {e}"
        )

        return None


def find_values_by_keys(obj, target_keys):
    """
    Recursively search a nested JSON object.

    Returns every value whose key matches one of
    target_keys.
    """

    results = []

    target_keys = {
        key.lower()
        for key in target_keys
    }

    def walk(value):

        if isinstance(value, dict):

            for key, child in value.items():

                if str(key).lower() in target_keys:
                    results.append(child)

                walk(child)

        elif isinstance(value, list):

            for item in value:
                walk(item)

    walk(obj)

    return results


def extract_description_from_initial_data(
    initial_data,
    job_id,
):
    """
    Extract a useful job description from Indeed's
    embedded _initialData.

    We deliberately search recursively because Indeed's
    internal JSON structure can change.
    """

    if not initial_data:
        return ""

    # --------------------------------------------------------
    # First, try fields that are most likely to contain
    # the actual job description.
    # --------------------------------------------------------

    candidate_keys = [
        "descriptionHtml",
        "jobDescriptionHtml",
        "descriptionText",
        "jobDescriptionText",
        "jobDescription",
        "description",
    ]

    candidates = find_values_by_keys(
        initial_data,
        candidate_keys,
    )

    best_description = ""

    for candidate in candidates:

        description = clean_description(
            candidate
        )

        # Ignore tiny strings such as:
        # "Job description" used as a label.
        if len(description) < 80:
            continue

        if len(description) > len(best_description):
            best_description = description

    if best_description:
        return best_description

    # --------------------------------------------------------
    # Fallback:
    # Search for larger text-bearing fields inside the
    # job-info area.
    # --------------------------------------------------------

    job_info_candidates = find_values_by_keys(
        initial_data,
        [
            "jobInfoModel",
            "jobInfoWrapperModel",
            "jobDescriptionSectionModel",
            "autoOpenTwoPaneViewjobResponse",
        ],
    )

    for section in job_info_candidates:

        nested_candidates = find_values_by_keys(
            section,
            candidate_keys,
        )

        for candidate in nested_candidates:

            description = clean_description(
                candidate
            )

            if len(description) < 80:
                continue

            if len(description) > len(best_description):
                best_description = description

    return best_description


# ============================================================
# SEARCH PAGE DESCRIPTION EXTRACTION
# ============================================================

def extract_description_from_search_page(
    page,
    job_id,
):
    """
    Extract the job description from Indeed's embedded
    search-page data.

    Important:
    We DO NOT open /viewjob because that endpoint was
    returning HTTP 403.

    Instead we use window._initialData already present
    on the successful search page.
    """

    initial_data = get_initial_data(
        page
    )

    if not initial_data:
        print(
            "        _initialData not available"
        )

        return ""

    description = (
        extract_description_from_initial_data(
            initial_data,
            job_id,
        )
    )

    if description:

        print(
            f"        Embedded description: "
            f"{len(description)} chars"
        )

        return description

    print(
        "        Embedded description not found"
    )

    return ""


# ============================================================
# MAIN INDEED SCRAPER
# ============================================================

def run_indeed(config):

    jobs = []

    queries = config["queries"]
    locations = config["locations"]

    results_per_search = config.get(
        "results_per_search",
        25,
    )

    headless = config.get(
        "headless",
        False,
    )

    # --------------------------------------------------------
    # Indeed "last 1 day" filter
    #
    # fromage=1 means jobs posted within the last day.
    # --------------------------------------------------------

    indeed_posted_within = config.get(
        "indeed_posted_within",
        1,
    )

    for query in queries:

        for location in locations:

            print(
                f"\n  Indeed: "
                f"{query} | {location}"
            )

            captured_for_search = 0

            # ------------------------------------------------
            # Fresh browser for every query/location.
            #
            # This matches the version that previously
            # worked better against Indeed's blocking.
            # ------------------------------------------------

            with sync_playwright() as p:

                browser = None
                context = None
                search_page = None

                try:

                    print(
                        "    Starting fresh browser..."
                    )

                    browser = p.chromium.launch(
                        headless=headless
                    )

                    context = browser.new_context(
                        viewport={
                            "width": 1440,
                            "height": 900,
                        },
                        locale="en-IN",
                        user_agent=(
                            "Mozilla/5.0 "
                            "(Macintosh; Intel Mac OS X 10_15_7) "
                            "AppleWebKit/537.36 "
                            "(KHTML, like Gecko) "
                            "Chrome/140.0.0.0 Safari/537.36"
                        ),
                    )

                    search_page = context.new_page()

                    # --------------------------------------------
                    # Only first result page.
                    #
                    # Pagination previously caused 403s.
                    # --------------------------------------------

                    url = (
                        f"{SEARCH_URL}"
                        f"?q={query}"
                        f"&l={location}"
                        f"&fromage={indeed_posted_within}"
                    )

                    print(
                        "    Opening first result page..."
                    )

                    try:

                        response = search_page.goto(
                            url,
                            wait_until="domcontentloaded",
                            timeout=30000,
                        )

                        search_page.wait_for_timeout(
                            2500
                        )

                        status = (
                            response.status
                            if response
                            else None
                        )

                        print(
                            f"    HTTP status: {status}"
                        )

                        page_title = clean_text(
                            search_page.title()
                        )

                        print(
                            f"    Page title: "
                            f"{page_title}"
                        )

                        # --------------------------------------------
                        # Detect access blocks.
                        # --------------------------------------------

                        title_lower = (
                            page_title.lower()
                        )

                        if (
                            status == 403
                            or "unusual traffic"
                            in title_lower
                            or "access denied"
                            in title_lower
                            or "just a moment"
                            in title_lower
                        ):

                            print(
                                "    Indeed access was blocked "
                                "for this search."
                            )

                            print(
                                "    Skipping this search."
                            )

                            continue

                        # --------------------------------------------
                        # Find cards.
                        # --------------------------------------------

                        cards = find_job_cards(
                            search_page
                        )

                        if cards is None:

                            print(
                                "    No job cards found."
                            )

                            continue

                        card_count = cards.count()

                        print(
                            f"    First page: "
                            f"{card_count} cards"
                        )

                        # --------------------------------------------
                        # Process cards.
                        # --------------------------------------------

                        for index in range(
                            card_count
                        ):

                            if (
                                captured_for_search
                                >= results_per_search
                            ):
                                break

                            card = cards.nth(
                                index
                            )

                            title = extract_title(
                                card
                            )

                            if not title:
                                continue

                            company = extract_company(
                                card
                            )

                            job_location = (
                                extract_location(
                                    card
                                )
                            )

                            salary = extract_salary(
                                card
                            )

                            link = extract_job_url(
                                card
                            )

                            job_id = extract_job_id(
                                link
                            )

                            # ------------------------------------
                            # Description:
                            #
                            # First try embedded _initialData.
                            # If unavailable, fall back to the
                            # visible card snippet.
                            # ------------------------------------

                            description = ""

                            if job_id:

                                description = (
                                    extract_description_from_search_page(
                                        search_page,
                                        job_id,
                                    )
                                )

                            if not description:

                                description = (
                                    extract_card_snippet(
                                        card
                                    )
                                )

                            print(
                                f"\n      {title}"
                            )

                            print(
                                f"        Company: "
                                f"{company}"
                            )

                            print(
                                f"        Location: "
                                f"{job_location}"
                            )

                            print(
                                f"        Salary: "
                                f"{salary}"
                            )

                            print(
                                f"        Job ID: "
                                f"{job_id}"
                            )

                            print(
                                f"        Description: "
                                f"{len(description)} chars"
                            )

                            print(
                                f"        URL: "
                                f"{link}"
                            )

                            # ------------------------------------
                            # Store normalized job.
                            # ------------------------------------

                            jobs.append({
                                "source": "Indeed",
                                "title": title,
                                "company": company,
                                "location": (
                                    job_location
                                    or location
                                ),
                                "experience": "",
                                "skills": "",
                                "description": description,
                                "url": link,
                            })

                            captured_for_search += 1

                            # Small delay between cards.
                            time.sleep(0.5)

                    except PlaywrightTimeoutError:

                        print(
                            "    Indeed first page timed out"
                        )

                    except Exception as e:

                        print(
                            f"    Indeed search failed: {e}"
                        )

                    print(
                        f"\n    captured: "
                        f"{captured_for_search}"
                    )

                    # Small pause before the next
                    # independent browser/search.
                    time.sleep(2)

                finally:

                    if search_page:

                        try:
                            search_page.close()
                        except Exception:
                            pass

                    if context:

                        try:
                            context.close()
                        except Exception:
                            pass

                    if browser:

                        try:
                            browser.close()
                        except Exception:
                            pass

    return jobs