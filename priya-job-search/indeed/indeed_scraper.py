import html
import json
import re
import time

from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


SEARCH_URL = "https://in.indeed.com/jobs"


def clean_text(text):
    if not text:
        return ""

    text = html.unescape(str(text))
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)

    return " ".join(text.split())


def extract_job_id(href):
    """
    Extract Indeed job ID from a job URL.

    Example:
        /viewjob?jk=abc123
    """

    if not href:
        return ""

    match = re.search(
        r"[?&]jk=([a-zA-Z0-9]+)",
        href
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


def get_text(card, selectors):
    """
    Try multiple selectors and return the first
    non-empty text.
    """

    for selector in selectors:
        try:
            element = card.locator(selector).first

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
    non-empty attribute value.
    """

    for selector in selectors:
        try:
            element = card.locator(selector).first

            if element.count() == 0:
                continue

            value = element.get_attribute(
                attribute,
                timeout=3000
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
            "h2.jobTitle",
            "a.jcs-JobTitle",
            "[data-testid='job-title']",
        ]
    )


def extract_company(card):
    return get_text(
        card,
        [
            "[data-testid='company-name']",
            ".companyName",
            "[class*='companyName']",
        ]
    )


def extract_location(card):
    return get_text(
        card,
        [
            "[data-testid='text-location']",
            ".companyLocation",
            "[class*='companyLocation']",
        ]
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
        ]
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
            "a[data-jk]",
            "a[href*='jk=']",
            "a[href*='/viewjob']",
        ],
        "href"
    )

    return make_absolute_url(href)


def extract_description(card):
    """
    Extract description/snippet directly from the search card
    when Indeed exposes one.
    """

    return get_text(
        card,
        [
            ".job-snippet",
            "[class*='job-snippet']",
            "[data-testid='job-snippet']",
        ]
    )


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
            locator = page.locator(selector)

            count = locator.count()

            print(
                f"    Selector "
                f"{selector}: "
                f"{count}"
            )

            if count > 0:
                return locator

        except Exception:
            continue

    return None


def find_description_in_json(obj):
    """
    Recursively search an Indeed embedded JSON object for
    job-description-like fields.

    We prioritize fields named jobDescription because Indeed's
    _initialData commonly stores the selected job description
    there.
    """

    exact_candidates = []
    fallback_candidates = []

    def walk(node):
        if isinstance(node, dict):

            for key, value in node.items():

                key_lower = str(key).lower()

                if isinstance(value, str):

                    text = clean_text(value)

                    if len(text) < 100:
                        continue

                    if key_lower in {
                        "jobdescription",
                        "jobdescriptionhtml",
                        "job_description",
                        "job_description_html",
                    }:
                        exact_candidates.append(text)

                    elif key_lower in {
                        "description",
                        "descriptiontext",
                        "descriptionhtml",
                    }:
                        fallback_candidates.append(text)

                walk(value)

        elif isinstance(node, list):

            for item in node:
                walk(item)

    walk(obj)

    if exact_candidates:
        return max(
            exact_candidates,
            key=len
        )

    if fallback_candidates:
        return max(
            fallback_candidates,
            key=len
        )

    return ""


def get_embedded_job_data(page):
    """
    Read Indeed's window._initialData from the current page.

    Returns:
        {
            "job_id": "...",
            "description": "..."
        }
    """

    try:
        data = page.evaluate(
            """
            () => {
                return window._initialData || null;
            }
            """
        )

        if not data:
            return {
                "job_id": "",
                "description": "",
            }

        selected_job_id = str(
            data.get("autoOpenTwoPaneJobKey")
            or data.get(
                "autoOpenJobAttributes",
                {}
            ).get("jobKey")
            or ""
        )

        response = data.get(
            "autoOpenTwoPaneViewjobResponse",
            {}
        )

        body = response.get(
            "body",
            {}
        )

        description = find_description_in_json(
            body
        )

        return {
            "job_id": selected_job_id,
            "description": description,
        }

    except Exception as e:
        print(
            f"        Embedded data extraction failed: {e}"
        )

        return {
            "job_id": "",
            "description": "",
        }


def extract_description_from_two_pane(
    page,
    expected_job_id,
):
    """
    Extract the full description from Indeed's two-pane
    job view.

    We deliberately verify the selected job ID before using
    embedded JSON so that one job's description is not copied
    to another job.
    """

    # ---------------------------------------------------------
    # 1. Try description DOM in the currently selected pane
    # ---------------------------------------------------------

    description_selectors = [
        "#jobDescriptionText",
        "[data-testid='jobDescriptionText']",
        ".jobsearch-JobComponent-description",
        "[class*='jobDescription']",
    ]

    for selector in description_selectors:

        try:
            locator = page.locator(
                selector
            ).first

            if locator.count() == 0:
                continue

            text = clean_text(
                locator.inner_text(
                    timeout=3000
                )
            )

            if len(text) >= 100:
                print(
                    f"        Description found in DOM "
                    f"using {selector} "
                    f"({len(text)} chars)"
                )

                return text

        except Exception:
            continue

    # ---------------------------------------------------------
    # 2. Try embedded window._initialData
    # ---------------------------------------------------------

    embedded = get_embedded_job_data(
        page
    )

    embedded_job_id = embedded["job_id"]
    embedded_description = embedded["description"]

    if embedded_job_id:
        print(
            f"        Embedded selected job ID: "
            f"{embedded_job_id}"
        )

    if expected_job_id:
        print(
            f"        Expected job ID: "
            f"{expected_job_id}"
        )

    # ---------------------------------------------------------
    # IMPORTANT:
    #
    # Only accept embedded description when the selected
    # job really is the job we are currently processing.
    # ---------------------------------------------------------

    if (
        expected_job_id
        and embedded_job_id
        and embedded_job_id != expected_job_id
    ):
        print(
            "        Embedded data belongs to a different "
            "job. Ignoring it."
        )

        return ""

    if embedded_description:
        print(
            f"        Embedded description found "
            f"({len(embedded_description)} chars)"
        )

        return embedded_description

    return ""


def click_job_card(page, card):
    """
    Click the job card/title so Indeed opens that job in
    the two-pane view.

    We intentionally click the search result rather than
    navigating directly to /viewjob because direct detail
    navigation was returning 403 in the current scraper.
    """

    selectors = [
        "a.jcs-JobTitle",
        "h2.jobTitle",
        "[data-testid='job-title']",
    ]

    for selector in selectors:

        try:
            locator = card.locator(
                selector
            ).first

            if locator.count() == 0:
                continue

            locator.scroll_into_view_if_needed(
                timeout=3000
            )

            locator.click(
                timeout=5000
            )

            return True

        except Exception:
            continue

    # Last attempt: click the card itself.

    try:
        card.scroll_into_view_if_needed(
            timeout=3000
        )

        card.click(
            timeout=5000
        )

        return True

    except Exception:
        return False


def wait_for_selected_job(
    page,
    expected_job_id,
    timeout_ms=7000,
):
    """
    Wait until Indeed's two-pane view appears to have
    selected the expected job.

    Returns True if the selected job ID matches.
    """

    start = time.time()

    while (
        (time.time() - start) * 1000
        < timeout_ms
    ):

        try:
            embedded = get_embedded_job_data(
                page
            )

            selected_job_id = embedded[
                "job_id"
            ]

            if (
                selected_job_id
                and selected_job_id == expected_job_id
            ):
                return True

        except Exception:
            pass

        # Also check whether the description DOM
        # has appeared.

        try:
            for selector in [
                "#jobDescriptionText",
                "[data-testid='jobDescriptionText']",
                ".jobsearch-JobComponent-description",
            ]:

                locator = page.locator(
                    selector
                ).first

                if locator.count() == 0:
                    continue

                text = clean_text(
                    locator.inner_text(
                        timeout=1000
                    )
                )

                if len(text) >= 100:
                    return True

        except Exception:
            pass

        page.wait_for_timeout(300)

    return False


def extract_full_description(
    page,
    card,
    job_id,
):
    """
    Click the selected search result and retrieve its
    full description.

    If the full description cannot be obtained, return the
    search-card snippet as fallback.
    """

    snippet = extract_description(
        card
    )

    if not job_id:
        return snippet

    print(
        "        Opening two-pane job view..."
    )

    clicked = click_job_card(
        page,
        card
    )

    if not clicked:
        print(
            "        Could not click job card."
        )

        return snippet

    # Give Indeed a moment to update the two-pane view.

    matched = wait_for_selected_job(
        page,
        job_id,
        timeout_ms=7000,
    )

    if not matched:
        print(
            "        Could not confirm selected job ID."
        )

    # Give the description a little additional time
    # because the two-pane content can load after the
    # selected job state changes.

    page.wait_for_timeout(
        1000
    )

    description = extract_description_from_two_pane(
        page,
        job_id,
    )

    if description:
        return description

    print(
        "        Full description unavailable; "
        "using search-card snippet."
    )

    return snippet


def run_indeed(config):
    jobs = []

    queries = config["queries"]
    locations = config["locations"]

    results_per_search = config.get(
        "results_per_search",
        25
    )

    headless = config.get(
        "headless",
        False
    )

    # Each query/location search gets its own fresh browser.
    #
    # This isolates individual searches so a block on one
    # request does not automatically terminate the whole run.
    #
    # This is not intended to bypass a CAPTCHA or other
    # access control.
    #
    # If a fresh browser receives a 403/block page, that search
    # is skipped and we continue with the remaining searches.

    for query in queries:

        for location in locations:

            print(
                f"\n  Indeed: "
                f"{query} | {location}"
            )

            captured_for_search = 0

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
                            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) "
                            "Chrome/140.0.0.0 Safari/537.36"
                        ),
                    )

                    search_page = context.new_page()

                    # -------------------------------------------------
                    # Open search page
                    # -------------------------------------------------

                    url = (
                        f"{SEARCH_URL}"
                        f"?q={query}"
                        f"&l={location}"
                    )

                    print(
                        "    Opening first result page..."
                    )

                    try:

                        response = search_page.goto(
                            url,
                            wait_until="domcontentloaded",
                            timeout=30000
                        )

                        search_page.wait_for_timeout(
                            2000
                        )

                        status = (
                            response.status
                            if response
                            else None
                        )

                        print(
                            f"    HTTP status: "
                            f"{status}"
                        )

                        page_title = clean_text(
                            search_page.title()
                        )

                        print(
                            f"    Page title: "
                            f"{page_title}"
                        )

                        # -------------------------------------------------
                        # Detect access-block pages
                        # -------------------------------------------------

                        if (
                            status == 403
                            or "unusual traffic"
                            in page_title.lower()
                            or "access denied"
                            in page_title.lower()
                            or "just a moment"
                            in page_title.lower()
                        ):

                            print(
                                "    Indeed access was blocked "
                                "for this search."
                            )

                            print(
                                "    Skipping this search and "
                                "continuing with the next one."
                            )

                            continue

                        # -------------------------------------------------
                        # Find job cards
                        # -------------------------------------------------

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

                        # -------------------------------------------------
                        # Process cards
                        # -------------------------------------------------

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
                                f"        URL: "
                                f"{link}"
                            )

                            # -------------------------------------------------
                            # IMPORTANT:
                            #
                            # Actually retrieve the description here.
                            # This is NOT just the old snippet assignment.
                            # -------------------------------------------------

                            description = (
                                extract_full_description(
                                    search_page,
                                    card,
                                    job_id,
                                )
                            )

                            print(
                                f"        Description: "
                                f"{len(description)} chars"
                            )

                            if description:
                                print(
                                    f"        Description preview: "
                                    f"{description[:300]}"
                                )
                            else:
                                print(
                                    "        Description: "
                                    "NOT FOUND"
                                )

                            # -------------------------------------------------
                            # Normalize to common job structure
                            # -------------------------------------------------

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

                            time.sleep(1)

                    except PlaywrightTimeoutError:

                        print(
                            "    Indeed search timed out."
                        )

                    except Exception as e:

                        print(
                            f"    Indeed search failed: {e}"
                        )

                finally:

                    if search_page:
                        search_page.close()

                    if context:
                        context.close()

                    if browser:
                        browser.close()

            print(
                f"\n    captured: "
                f"{captured_for_search}"
            )

    print(
        f"\n  Indeed total captured: "
        f"{len(jobs)} jobs"
    )

    return jobs