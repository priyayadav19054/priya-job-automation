import asyncio
from pathlib import Path
from playwright.async_api import async_playwright
from src.excel import get_next_pending_job

async def main():
    job = get_next_pending_job()
    if not job:
        print("No PENDING jobs found.")
        return
    print(f"Processing: {job.get('Company')} - {job.get('Title')}")
    print(f"URL: {job.get('URL')}")
    screenshot_dir = Path("screenshots") / str(job.get("Rank", "unknown"))
    screenshot_dir.mkdir(parents=True, exist_ok=True)
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=100)
        context = await browser.new_context()
        page = await context.new_page()
        await page.goto(job["URL"], wait_until="domcontentloaded", timeout=60000)
        await page.screenshot(path=str(screenshot_dir / "01_opened.png"), full_page=True)
        print("V1 stops here — nothing is filled or submitted.")
        input("Press ENTER to close the browser...")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
