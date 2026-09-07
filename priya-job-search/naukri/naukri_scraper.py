from playwright.sync_api import sync_playwright
import re


def run_naukri(config):
    jobs = []

    queries = config["queries"]
    locations = config["locations"]

    results_per_search = config.get("results_per_search", 25)

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=config.get("headless", False)
        )

        search_page = browser.new_page()
        detail_page = browser.new_page()

        for query in queries:
            for location in locations:

                print(f"  Naukri: {query} | {location}")

                try:
                    slug_q = re.sub(
                        r"[^a-z0-9]+",
                        "-",
                        query.lower()
                    ).strip("-")

                    slug_l = re.sub(
                        r"[^a-z0-9]+",
                        "-",
                        location.lower()
                    ).strip("-")

                    # jobAge=1 = jobs posted within the last 1 day
                    url = (
                        f"https://www.naukri.com/"
                        f"{slug_q}-jobs-in-{slug_l}"
                        f"?jobAge=1"
                    )

                    response = search_page.goto(
                        url,
                        wait_until="domcontentloaded",
                        timeout=30000
                    )

                    search_page.wait_for_timeout(2500)

                    if response and response.status >= 400:
                        print(
                            f"    HTTP {response.status}"
                        )
                        continue

                    selectors = [
                        "div.srp-jobtuple-wrapper",
                        "div.cust-job-tuple",
                        "article.jobTuple",
                        "div[class*='jobTuple']",
                        "div[class*='cust-job-tuple']"
                    ]

                    cards = []

                    for selector in selectors:

                        count = search_page.locator(
                            selector
                        ).count()

                        if count:
                            cards = search_page.locator(
                                selector
                            ).all()
                            break

                    print(f"    Search cards: {len(cards)}")

                    for card in cards[:results_per_search]:

                        def txt(selector):
                            locator = card.locator(selector)

                            if locator.count():
                                return locator.first.inner_text().strip()

                            return ""

                        title = txt("a[class*='title']")
                        company = txt("a[class*='comp-name']")

                        loc = (
                            txt("span[class*='locWdth']")
                            or location
                        )

                        exp = txt("span[class*='expwdth']")

                        # Try several possible selectors for posting age
                        posted = ""

                        posted_selectors = [
                            "span.job-post-day",
                            "span[class*='job-post-day']",
                            "span[class*='posted']",
                            "span[class*='date']",
                            ".job-post-day"
                        ]

                        for posted_selector in posted_selectors:
                            value = txt(posted_selector)

                            if value:
                                posted = value
                                break

                        link = ""

                        title_locator = card.locator(
                            "a[class*='title']"
                        )

                        if title_locator.count():
                            link = (
                                title_locator
                                .first
                                .get_attribute("href")
                                or ""
                            )

                        if not title:
                            continue

                        print(
                            f"    Job: {title}"
                            + (
                                f" | Posted: {posted}"
                                if posted
                                else ""
                            )
                        )

                        description = ""

                        if link:

                            try:
                                print(
                                    f"      Fetching JD..."
                                )

                                detail_page.goto(
                                    link,
                                    wait_until="domcontentloaded",
                                    timeout=30000
                                )

                                detail_page.wait_for_timeout(1500)

                                # Naukri has used several structures
                                # for the JD container.
                                description_selectors = [
                                    ".job-desc .dang-inner-html",
                                    ".dang-inner-html",
                                    ".job-desc",
                                    "section.job-desc",
                                    "[class*='job-desc']"
                                ]

                                for selector in description_selectors:

                                    locator = detail_page.locator(
                                        selector
                                    )

                                    if locator.count():

                                        text = (
                                            locator
                                            .first
                                            .inner_text()
                                            .strip()
                                        )

                                        # Avoid accidentally capturing
                                        # a tiny unrelated element.
                                        if len(text) >= 50:
                                            description = text
                                            break

                                # Final fallback:
                                # locate the "Job Description" heading
                                # and inspect its nearby container.
                                if not description:

                                    heading = detail_page.get_by_text(
                                        "Job Description",
                                        exact=True
                                    )

                                    if heading.count():

                                        parent = heading.first.locator(
                                            ".."
                                        )

                                        text = (
                                            parent
                                            .inner_text()
                                            .strip()
                                        )

                                        if len(text) >= 50:
                                            description = text

                                if description:
                                    print(
                                        f"      JD captured: "
                                        f"{len(description)} chars"
                                    )
                                else:
                                    print(
                                        "      JD not available"
                                    )

                            except Exception as e:
                                print(
                                    f"      JD fetch failed: {e}"
                                )

                        jobs.append({
                            "source": "Naukri",
                            "title": title,
                            "company": company,
                            "location": loc,
                            "experience": exp,
                            "skills": "",
                            "description": description,
                            "url": link or "",
                            "posted": posted
                        })

                except Exception as e:
                    print(f"    skipped: {e}")

        browser.close()

    return jobs