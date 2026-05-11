from openpyxl.worksheet import page
import re
from playwright.sync_api import Playwright, sync_playwright

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    
    page.goto("https://watermelon-life-insurance.vercel.app/fluttertesting/index.html")
    page.wait_for_timeout(3000)  # wait for Flutter to load
    
    # Click the flutter view to focus it first
    page.locator("flutter-view").click()
    page.wait_for_timeout(500)

    # Type username
    page.keyboard.type("flutteruser")
    page.wait_for_timeout(300)

    # Tab to password field
    page.keyboard.press("Tab")
    page.wait_for_timeout(300)

    # Type password
    page.keyboard.type("password")
    page.wait_for_timeout(300)

    # Tab to login button
    page.keyboard.press("Tab")
    page.wait_for_timeout(300)

    # Press Enter to click login
    page.keyboard.press("Enter")
    page.wait_for_timeout(3000)  # wait for OTP screen to load

    page.keyboard.press("Tab")
    # OTP screen — type each digit and Tab to next
    page.keyboard.type("1")
    # page.keyboard.press("Tab")
    page.wait_for_timeout(200)

    page.keyboard.type("2")
    # page.keyboard.press("Tab")
    page.wait_for_timeout(200)

    page.keyboard.type("3")
    # page.keyboard.press("Tab")
    page.wait_for_timeout(200)

    page.keyboard.type("4")
    # page.keyboard.press("Tab")
    page.wait_for_timeout(200)

    page.keyboard.type("5")
    # page.keyboard.press("Tab")
    page.wait_for_timeout(200)

    page.keyboard.type("6")
    page.wait_for_timeout(300)

    page.keyboard.press("Tab")
    # Submit OTP
    page.keyboard.press("Enter")
    page.wait_for_timeout(2000)

    with page.expect_file_chooser() as fc_info:
        page.keyboard.press("Tab")
        page.keyboard.press("Tab")
        page.keyboard.press("Enter")  # this opens the file dialog

    file_chooser = fc_info.value
    file_chooser.set_files("/Users/rounaknaik/Downloads/excel for flutter testing.xlsx")

    page.wait_for_timeout(10000)
    # # Click Upload Excel button and handle file chooser
    # with page.expect_file_chooser() as fc_info:
    #     page.keyboard.press("Enter")  # since you're using Tab to navigate to the button
    # file_chooser = fc_info.value
    # file_chooser.set_files("/Users/rounaknaik/Downloads/excel for flutter testing.xlsx")
    # page.locator("input[type='file']").set_input_files("/Users/rounaknaik/Downloads/excel for flutter testing.xlsx")

    # Wait 10 seconds after upload
    # page.wait_for_timeout(10000)

    # # Wait 10 seconds after upload
    # page.wait_for_timeout(10000)

    context.close()
    browser.close()

with sync_playwright() as playwright:
    run(playwright)