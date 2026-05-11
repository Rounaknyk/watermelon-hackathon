# axis_maxlife_premium_calc.py
# ------------------------------------------------------------
# Playwright‑Python script that automates the full quote flow
# on https://www.axismaxlife.com/term-insurance-plans/premium-calculator
# ------------------------------------------------------------

import sys
from playwright.sync_api import sync_playwright, expect

# ----------------------------------------------------------------
# Helper: wait for an element and fill it (covers hidden/auto‑complete)
# ----------------------------------------------------------------
def fill_input(page, selector, value):
    """Robustly fill a text input or textarea."""
    page.wait_for_selector(selector, timeout=15000)
    page.locator(selector).fill(value)


def click_button(page, selector):
    """Click a button/link; wait for it to be enabled."""
    page.wait_for_selector(selector, timeout=15000)
    page.locator(selector).click()


def run():
    with sync_playwright() as p:
        # ----------------------------------------------------------------
        # 1️⃣ Launch a visible browser (headless=False) – you can set
        #    headless=True for CI runs.
        # ----------------------------------------------------------------
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # ----------------------------------------------------------------
        # 2️⃣ Navigate to the calculator
        # ----------------------------------------------------------------
        page.goto(
            "https://www.axismaxlife.com/term-insurance-plans/premium-calculator",
            wait_until="load",
            timeout=30000,
        )

        # ----------------------------------------------------------------
        # 3️⃣ Fill the form fields
        # ----------------------------------------------------------------
        # Full Name
        fill_input(page, 'input[name="fullName"]', "Vikram Banerjee")

        # Date of Birth – the control expects DD/MM/YYYY
        fill_input(page, 'input[name="dob"]', "19/08/1982")

        # # Mobile number (includes country code selector)
        # # First select country code ‘+91 India’
        # click_button(page, 'button[data-testid="country-code"]')          # open dropdown
        # click_button(page, 'li[data-value="+91"]')                        # choose +91
        fill_input(page, '#mobile', "9069041652")

        # Annual Income (numeric field)
        
        fill_input(page, 'input[name="annualIncome"]', "3812052")

        # Gender – choose Male
        click_button(page, 'label[for="genderMale"]')

        # Occupation – Self Employed (radio/checkbox)
        click_button(page, 'label[for="occupationSelfEmployed"]')

        # Education – Grad or above
        click_button(page, 'label[for="educationGrad"]')

        # Life Cover – 3 crore (dropdown) 
        click_button(page, 'button[data-testid="life-cover"]')
        click_button(page, 'li[data-value="30000000"]')   # 3 crore = 30,000,000

        # Cover Till Age – 65
        click_button(page, 'button[data-testid="cover-till-age"]')
        click_button(page, 'li[data-value="65"]')

        # Critical Illness Rider – No
        click_button(page, 'label[for="criticalIllnessNo"]')

        # Email
        fill_input(page, 'input[name="email"]', "vikram.banerjee@gmail.com")

        # Pincode – auto‑fills city
        fill_input(page, 'input[name="pincode"]', "560016")
        # Wait for city field to be populated (a short pause is enough)
        page.wait_for_timeout(2000)

        # NRI – No (checkbox/radio)
        click_button(page, 'label[for="nriNo"]')

        # City field (auto‑filled) – just verify it contains “Bangalore”
        city_input = page.locator('input[name="city"]')
        city_value = city_input.input_value()
        if "Bangalore" not in city_value:
            print(f"⚠️  Expected city Bangalore but got: {city_value}")

        # Diabetes – No
        click_button(page, 'label[for="diabetesNo"]')

        # Marital Status – Single
        click_button(page, 'label[for="maritalSingle"]')

        # ------------------------------------------------------------
        # 4️⃣ Submit – the button that triggers the quote popup
        # ------------------------------------------------------------
        # The button often has text like “Proceed”, “Get Quote”, or “Calculate”.
        # We first try a direct text selector, then fall back to a generic
        # role‑button selector.
        try:
            click_button(page, 'text="Proceed"')
        except Exception:
            try:
                click_button(page, 'text="Get Quote"')
            except Exception:
                click_button(page, 'button[type="submit"]')

        # ----------------------------------------------------------------
        # 5️⃣ Wait for the popup / result page and verify a quote appears
        # ----------------------------------------------------------------
        # The result opens in a new window (popup). Capture it.
        with page.expect_popup() as popup_info:
            # the click above already triggered the popup – just ensure we have it
            pass
        popup = popup_info.value

        # Wait for a reasonable element that shows the premium amount
        popup.wait_for_selector('text="Your Quote"', timeout=15000)
        premium_text = popup.locator('div[data-testid="premium-amount"]').inner_text()
        print("\n✅ Quote displayed:", premium_text)

        # ----------------------------------------------------------------
        # 6️⃣ Clean up
        # ----------------------------------------------------------------
        popup.close()
        context.close()
        browser.close()


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print("\n❌ Test failed:", e)
        sys.exit(1)
