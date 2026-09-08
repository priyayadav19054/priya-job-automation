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


def extract_description(card):
    """
    Indeed currently does not appear to expose a job
    description/snippet in the search card HTML.

    Keep this function so the normalized job structure
    remains compatible with the common pipeline.
    """

    return get_text(
        card,
        [
            ".job-snippet",
            "[class*='job-snippet']",
            "[data-testid='job-snippet']",
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
                f"    Selector "
                f"{selector}: "
                f"{count}"
            )

            if count > 0:
                return locator

        except Exception:
            continue

    return None


def run_indeed(config):

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

    headless = config.get(
        "headless",
        False
    )

    with sync_playwright() as p:

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

        for query in queries:

            for location in locations:

                print(
                    f"\n  Indeed: "
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

                    url = (
                        f"{SEARCH_URL}"
                        f"?q={query}"
                        f"&l={location}"
                        f"&start={start}"
                    )

                    print(
                        f"    Opening page "
                        f"{page_number + 1}..."
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

                        # ---------------------------------------------
                        # Detect obvious access-block pages
                        # ---------------------------------------------

                        if (
                            status == 403
                            or "unusual traffic"
                            in page_title.lower()
                            or "access denied"
                            in page_title.lower()
                        ):

                            print(
                                "    Indeed access "
                                "was blocked."
                            )

                            break

                        # ---------------------------------------------
                        # Find job cards
                        # ---------------------------------------------

                        cards = find_job_cards(
                            search_page
                        )

                        if cards is None:

                            print(
                                "    No job cards "
                                "found."
                            )

                            break

                        card_count = cards.count()

                        print(
                            f"    Page "
                            f"{page_number + 1}: "
                            f"{card_count} cards"
                        )

                        # ---------------------------------------------
                        # Process cards
                        # ---------------------------------------------

                        for index in range(
                            card_count
                        ):

                            if (
                                captured_for_search
                                >= results_per_search
                            ):
                                break

                            card = cards.nth(index)

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

                            snippet = (
                                extract_description(
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
                                f"        URL: "
                                f"{link}"
                            )

                            print(
                                f"        Snippet: "
                                f"{snippet[:200]}"
                            )

                            # -----------------------------------------
                            # Description
                            #
                            # Indeed's current search card does not
                            # expose the description, so use the
                            # snippet if one is available.
                            # -----------------------------------------

                            description = snippet

                            # -----------------------------------------
                            # Normalize to common job structure
                            # -----------------------------------------

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
                            f"    Page "
                            f"{page_number + 1} "
                            f"timed out"
                        )

                    except Exception as e:

                        print(
                            f"    Indeed page "
                            f"{page_number + 1} "
                            f"failed: {e}"
                        )

                    time.sleep(2)

                print(
                    f"\n    captured: "
                    f"{captured_for_search}"
                )

        search_page.close()

        context.close()
        browser.close()

    return jobs