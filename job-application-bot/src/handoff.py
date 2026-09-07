async def human_handoff(page, reason):
    print("\n" + "="*60); print("HUMAN ACTION REQUIRED"); print(f"Reason: {reason}"); print("Complete the required action in the visible browser."); print("Do not attempt to bypass CAPTCHA or verification."); print("="*60)
    input("Press ENTER when finished...")
    return True
