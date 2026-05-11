import re
import csv
import uuid
import openpyxl
from playwright.sync_api import Playwright, sync_playwright

# ── CONFIG ──────────────────────────────────────────────
EXCEL_FILE = "insurance_test_data.xlsx"
OUTPUT_CSV = "quote_results.csv"
MAX_USERS  = 5   # change this to run more

# ── DEFAULTS (used when value missing/unknown) ───────────
DEFAULTS = {
    "gender":     "Male",
    "tobacco":    "No",
    "language":   "English",
    "occupation": "Salaried",
    "education":  "Graduate & Above",
    "diabetic":   "No",
    "marital":    "Single",
    "life_cover": "50 Lakhs",
    "cover_age":  "60",
    "city":       "Bangalore",
}

# ── MAPPINGS (Excel value → what the site shows) ─────────
EDUCATION_MAP = {
    "grad or above": "Graduate & Above",
    "graduate & above": "Graduate & Above",
    "12th pass":     "12",
    "10th pass":     "10",
}

OCCUPATION_MAP = {
    "salaried":      "Salaried",
    "self employed": "Self-employed/Business",
    "self-employed": "Self-employed/Business",
    "housewife":     "Housewife",   # no housewife option on UI — default to Salaried
}

# ── READ EXCEL ───────────────────────────────────────────
def load_users(filepath, max_users):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() for h in rows[0]]
    users = []
    for row in rows[1 : max_users + 1]:
        user = dict(zip(headers, row))
        users.append(user)
    return users

# ── HELPERS ──────────────────────────────────────────────
def get(user, key, default=""):
    """Safe getter — returns default if None or missing."""
    val = user.get(key)
    return str(val).strip() if val is not None else default

def map_education(val):
    return EDUCATION_MAP.get(val.lower(), DEFAULTS["education"])

def map_occupation(val):
    return OCCUPATION_MAP.get(val.lower(), DEFAULTS["occupation"])

def normalize_cover(val):
    """Convert '3 crore', '1.5 crore', '50 lakhs' → display label fragment."""
    val = val.lower().strip()
    if "crore" in val:
        num = val.replace("crore", "").strip()
        return f"{num} Crore" if num != "1" else "1 Crore"
    if "lakh" in val or "lac" in val:
        num = re.search(r"[\d.]+", val).group()
        return f"{num} Lakh"
    return val

# ── SAVE RESULT ──────────────────────────────────────────
def save_to_csv(row_data):
    fieldnames = ["Test ID", "Insurer Name", "Equote Number", "Premium (1st Year)"]
    file_exists = True
    try:
        open(OUTPUT_CSV, "r")
    except FileNotFoundError:
        file_exists = False

    with open(OUTPUT_CSV, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row_data)

# ── EXTRACT RESULTS FROM PAGE ────────────────────────────
def extract_results(page, insurer_name):
    page.wait_for_selector("text=Review your Term Insurance summary", timeout=15000)
    page.wait_for_timeout(2000)

    summary_text = page.inner_text("body")
    lines = [l.strip() for l in summary_text.split("\n") if l.strip()]

    # Equote — line immediately after "Equote Number"
    equote_number = "N/A"
    try:
        idx = next(i for i, l in enumerate(lines) if "Equote Number" in l)
        equote_number = lines[idx + 1].strip()
    except (StopIteration, IndexError):
        pass

    # Premium — ₹ value after "Premium for 1st year"
    premium = "N/A"
    try:
        idx = next(i for i, l in enumerate(lines) if "Premium for 1st" in l)
        for l in lines[idx:idx + 6]:
            match = re.search(r"₹\s*([\d,]+\.?\d*)", l)
            if match:
                premium = match.group(1).replace(",", "")
                break
    except StopIteration:
        pass

    return {
        "Test ID":           str(uuid.uuid4())[:8].upper(),
        "Insurer Name":      insurer_name,
        "Equote Number":     equote_number,
        "Premium (1st Year)": premium,
    }

# ── MAIN FLOW PER USER ───────────────────────────────────
def run_for_user(playwright, user):
    name       = get(user, "Full Name",     "Test User")
    dob        = get(user, "Date of birth", "01/01/1990")
    mobile     = get(user, "Mobile",        "9999999999")
    email      = get(user, "email id",      "test@gmail.com")
    income     = get(user, "Annual income", "1000000").replace(",", "")
    pincode    = get(user, "Pincode",       "560001")
    gender     = get(user, "Gender",        DEFAULTS["gender"])
    occupation = map_occupation(get(user, "Occupation", DEFAULTS["occupation"]))
    education  = map_education(get(user, "Education",  DEFAULTS["education"]))
    life_cover = normalize_cover(get(user, "Life cover", DEFAULTS["life_cover"]))
    cover_age  = str(get(user, "Cover till age", DEFAULTS["cover_age"]))
    rider      = get(user, "Critical Illness Rider", "No")

    print(f"\n{'='*50}")
    print(f"Running for: {name}")
    print(f"  DOB: {dob} | Mobile: {mobile} | Gender: {gender}")
    print(f"  Occupation: {occupation} | Education: {education}")
    print(f"  Life Cover: {life_cover} | Cover Age: {cover_age} yrs")
    print(f"  Rider: {rider} | Pincode: {pincode}")
    print(f"{'='*50}")

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page    = context.new_page()

    try:
        # ── STEP 1: Personal details ────────────────────
        page.goto("https://www.axismaxlife.com/term-insurance-plans/premium-calculator")
        page.get_by_role("textbox", name="Full Name*").fill(name)
        page.get_by_role("textbox", name="Date of Birth*").fill(dob)
        page.locator("#mobile").fill(mobile)
        page.locator("label").filter(has_text="- 7").click()
        page.get_by_role("button", name="button").click()
        page.wait_for_timeout(3000)

        # ── STEP 2: Profile questions ───────────────────
        page.get_by_text(gender, exact=True).click()
        page.get_by_text("No", exact=True).click()       # tobacco
        page.get_by_text("English", exact=True).click()  # language
        page.get_by_text("Salaried", exact=True).click()
        page.get_by_text("Graduate & Above", exact=False).click()
        page.get_by_text("No", exact=True).click()       # diabetic
        page.get_by_text("Single", exact=True).click()   # marital
        page.get_by_text("Check Coverage", exact=True).click()

        # ── STEP 3: Customize plan ──────────────────────
        page.wait_for_selector("text=customize your Term Plan")
        page.wait_for_timeout(2000)

        # Open life cover dropdown
        page.get_by_text(re.compile(r"Recommended|Market Linked Returns")).first.click()
        page.wait_for_timeout(1000)

        # Select life cover by partial match
        # Select life cover by partial match (with fallback to default if not found)
        try:
            target_cover = page.locator("label, li").filter(has_text=re.compile(life_cover, re.IGNORECASE)).first
            target_cover.click(timeout=5000)
        except Exception:
            default_label = normalize_cover(DEFAULTS["life_cover"])
            print(f"  ⚠ Life cover '{life_cover}' not found, choosing default: {default_label}")
            page.locator("label, li").filter(has_text=re.compile(default_label, re.IGNORECASE)).first.click()
        page.wait_for_timeout(1000)

        # # Select cover till age
        # page.get_by_text(f"{cover_age} years", exact=False).first.click()
        # page.wait_for_timeout(1000)


        page.get_by_text("MORE", exact=True).first.click()
        page.get_by_role("textbox", name="Enter your custom value").fill(cover_age)
        
        # page.get(cover_age, exact=True).first.click()
        # page.get_by_text(f"{cover_age} years", exact=False).first.click()
        # page.wait_for_timeout(1000)
        page.get_by_text("Confirm", exact=True).first.click() 

        page.get_by_text("Proceed", exact=False).first.click()
        page.wait_for_timeout(1000)

        # ── Critical Illness Rider ──────────────────────
        if rider.lower() == "yes":
            try:
                page.get_by_text("Critical Illness", exact=False).first.click()
                page.wait_for_timeout(500)
            except Exception:
                print("  ⚠ Could not click Critical Illness rider — skipping")

        try:
            page.locator("text=/Skip|Proceed/").first.click(timeout=5000)
            page.wait_for_timeout(1000)
        except Exception:
            # If neither is found, we assume the flow moved forward or the button isn't needed
            pass


        # ── STEP 3b: Eligibility details ───────────────
        page.get_by_role("textbox", name="Email Address*").fill(email)
        page.get_by_role("textbox", name="Annual Income*").fill(income)
        page.get_by_role("textbox", name="Pincode of Current Residential Address*").fill(pincode)
        page.wait_for_timeout(2000)

        # City dropdown
        page.locator("button.city-select").click()
        page.wait_for_timeout(500)
        page.get_by_text("Bangalore", exact=True).first.click()

        page.get_by_text("Download Benefit Illustration").click()
        page.wait_for_timeout(8000)
        page.get_by_text("Proceed", exact=True).first.click()
        page.wait_for_timeout(5000)

        # ── STEP 4: Extract & save ──────────────────────
        result = extract_results(page, name)
        save_to_csv(result)

        print(f"  ✓ Test ID:  {result['Test ID']}")
        print(f"  ✓ Equote:   {result['Equote Number']}")
        print(f"  ✓ Premium:  ₹{result['Premium (1st Year)']}")
        print(f"  ✓ Saved to {OUTPUT_CSV}")

    except Exception as e:
        print(f"  ✗ ERROR for {name}: {e}")
        # Still save a row so we know it failed
        save_to_csv({
            "Test ID":            str(uuid.uuid4())[:8].upper(),
            "Insurer Name":       name,
            "Equote Number":      "ERROR",
            "Premium (1st Year)": str(e)[:60],
        })
    finally:
        page.wait_for_timeout(2000)
        context.close()
        browser.close()

# ── ENTRY POINT ──────────────────────────────────────────
def main():
    users = load_users(EXCEL_FILE, MAX_USERS)
    print(f"Loaded {len(users)} users from {EXCEL_FILE}")

    with sync_playwright() as playwright:
        for i, user in enumerate(users, 1):
            print(f"\n[{i}/{len(users)}] Starting...")
            run_for_user(playwright, user)

    print(f"\n✅ Done! Results saved to {OUTPUT_CSV}")

if __name__ == "__main__":
    main()