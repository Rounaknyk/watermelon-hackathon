import re
import csv
import uuid
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.axismaxlife.com/term-insurance-plans/premium-calculator")
    page.get_by_role("textbox", name="Full Name*").click()
    page.get_by_role("textbox", name="Full Name*").press("CapsLock")
    page.get_by_role("textbox", name="Full Name*").fill("Rounak ")
    page.get_by_role("textbox", name="Full Name*").press("CapsLock")
    page.get_by_role("textbox", name="Full Name*").fill("Rounak Naik")
    page.get_by_role("textbox", name="Date of Birth*").click()
    page.get_by_role("textbox", name="Date of Birth*").fill("19/06/2004")
    page.locator("#mobile").click()
    page.locator("#mobile").fill("9322946253")
    page.locator("label").filter(has_text="- 7").click()
    page.get_by_role("button", name="button").click()
    # wait for popup/modal
    page.wait_for_timeout(3000)

    # Gender
    page.get_by_text("Female", exact=True).click()

    # Tobacco/Nicotine
    page.get_by_text("No", exact=True).click()

    # Preferred Language
    page.get_by_text("English", exact=True).click()

    # Occupation
    page.get_by_text("Salaried", exact=True).click()

    # Education
    page.get_by_text("Graduate & Above").click()

    # Are you diabetic
    page.get_by_text("No", exact=True).click()

    # Marital status
    page.get_by_text("Single", exact=True).click()

    page.get_by_text("Check Coverage", exact=True).click()

    # STEP 3
    page.wait_for_selector("text=customize your Term Plan")
    page.wait_for_timeout(2000)

    page.get_by_text("Recommended").click()
    page.wait_for_timeout(1000)

    page.locator("label, li").filter(has_text=re.compile(r"50 Lakhs")).click()
    page.wait_for_timeout(1000)

    page.get_by_text("65 years").click()
    page.wait_for_timeout(1000)

    page.get_by_text("Proceed", exact=False).click()
    page.wait_for_timeout(1000)

    page.get_by_text("Critical Illness & Disability Rider", exact=False).click()

    page.get_by_text("Skip").click()

    page.get_by_role("textbox", name="Email Address*").fill("rounak@gmail.com")
    page.get_by_role("textbox", name="Annual Income*").fill("2000000")
    page.get_by_role("textbox", name="Pincode of Current Residential Address*").fill("403802")
    page.wait_for_timeout(2000)
    page.locator("button.city-select").click()
    page.wait_for_timeout(500)
    page.get_by_text("Bangalore", exact=True).first.click()

    page.get_by_text("Download Benefit Illustration").click()

    page.wait_for_timeout(8000)
    page.get_by_text("Proceed", exact=True).first.click()

    page.wait_for_timeout(5000)

    # # ── STEP 4 — Extract quote data and save to CSV ──
    # page.wait_for_selector("text=Review your Term Insurance summary")
    # page.wait_for_timeout(2000)

    # Extract insurer name
    insurer_name = "rounak"
    # insurer_name = page.locator("h2, h3").filter(
    #     has_text=re.compile(r"Hi (.+),")
    # ).first.inner_text()
    # insurer_name = re.sub(r"Hi\s*|,", "", insurer_name).strip()

    # Extract equote number
    equote_row = page.locator("tr, div, p").filter(
        has_text=re.compile(r"^\d{4,}[A-Z]+$")  # matches codes like 5005XXYQ
    ).first.inner_text().strip()
    equote_number = equote_row

    # Extract premium for 1st year
    premium_element = page.locator(
        "xpath=//p[contains(text(),'Premium for 1')]/following-sibling::p[1] | "
        "//span[contains(text(),'Premium for 1')]/following-sibling::span[1]"
    ).first.inner_text()
    premium = re.search(r"[\d,]+\.?\d*", premium_element).group().replace(",", "")

    # Generate unique test ID
    test_id = str(uuid.uuid4())[:8].upper()

    print(f"Insurer:  {insurer_name}")
    print(f"Test ID:  {test_id}")
    print(f"Equote:   {equote_number}")
    print(f"Premium:  {premium}")

    # Write to CSV
    csv_file = "quote_results.csv"
    file_exists = False
    try:
        open(csv_file, "r")
        file_exists = True
    except FileNotFoundError:
        pass

    with open(csv_file, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "Test ID", "Insurer Name", "Equote Number", "Premium (1st Year)"
        ])
        if not file_exists:
            writer.writeheader()   # write header only once
        writer.writerow({
            "Test ID": test_id,
            "Insurer Name": insurer_name,
            "Equote Number": equote_number,
            "Premium (1st Year)": premium
        })

    print(f"Saved to {csv_file}")

    page.wait_for_timeout(3000)
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)