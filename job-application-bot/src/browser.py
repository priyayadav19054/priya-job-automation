from playwright.async_api import async_playwright
class BrowserManager:
    def __init__(self, headless=False, slow_mo=100): self.headless=headless; self.slow_mo=slow_mo; self.playwright=None; self.browser=None; self.context=None
    async def start(self):
        self.playwright=await async_playwright().start(); self.browser=await self.playwright.chromium.launch(headless=self.headless, slow_mo=self.slow_mo); self.context=await self.browser.new_context(); return await self.context.new_page()
    async def close(self):
        if self.browser: await self.browser.close()
        if self.playwright: await self.playwright.stop()
