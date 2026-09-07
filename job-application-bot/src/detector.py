CAPTCHA_TERMS=["captcha","recaptcha","hcaptcha","verify you are human","cloudflare challenge"]
OTP_TERMS=["verification code","one-time password","otp","security code"]
async def page_contains_terms(page, terms):
    return any(t in (await page.locator("body").inner_text()).lower() for t in terms)
async def detect_captcha(page): return await page_contains_terms(page, CAPTCHA_TERMS)
async def detect_otp(page): return await page_contains_terms(page, OTP_TERMS)
