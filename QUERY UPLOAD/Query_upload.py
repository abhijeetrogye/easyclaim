import time
import logging
import os
from functools import wraps
import pandas as pd
from playwright.sync_api import sync_playwright, Playwright

# ---------------- CONFIG ---------------- #
USERNAME = os.getenv("NHA_USERNAME", "dineshdwarka")
HEADLESS = False
TIMEOUT = 45_000
EXCEL_FILE = "query_upload.xlsx"
MAX_RETRIES = 3
RETRY_DELAY = 2

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ---------------- RETRY DECORATOR ---------------- #
def retry_on_failure(max_retries=MAX_RETRIES, delay=RETRY_DELAY):
    """Decorator to retry a function on failure with exponential backoff."""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_exception = e
                    if attempt < max_retries - 1:
                        wait_time = delay * (2 ** attempt)
                        logging.warning(f"Attempt {attempt + 1} failed: {e}. Retrying in {wait_time}s...")
                        time.sleep(wait_time)
                    else:
                        logging.error(f"All {max_retries} attempts failed for {func.__name__}")
            raise last_exception
        return wrapper
    return decorator

# ---------------- UTILITIES ---------------- #
@retry_on_failure()
def wait_and_click(locator, name="element"):
    logging.info(f"Waiting for and clicking: {name}")
    locator.wait_for(state="visible", timeout=TIMEOUT)
    locator.click()

@retry_on_failure()
def wait_and_fill(locator, value, name="field"):
    locator.wait_for(state="visible", timeout=TIMEOUT)
    clean_value = "" if pd.isna(value) or str(value).lower() == "nan" else str(value)
    logging.info(f"Filling {name} with: '{clean_value}'")
    locator.fill(clean_value)

def wait_for_network_idle(page, timeout=5000):
    """Wait for network to be idle."""
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except:
        logging.warning("Network idle timeout - continuing anyway")

# ---------------- EXCEL ---------------- #
def load_excel():
    """Load Excel file with error handling."""
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")

    df = pd.read_excel(EXCEL_FILE, dtype={"registeration_id": str, "registration_id": str})

    if "registration_id" in df.columns and "registeration_id" not in df.columns:
        df.rename(columns={"registration_id": "registeration_id"}, inplace=True)

    for col in ["remarks", "status", "supporting_doc"]:
        if col not in df.columns:
            df[col] = ""

    df["remarks"] = df["remarks"].fillna("")
    df["status"] = df["status"].fillna("")

    return df

def save_excel(df, max_retries=3):
    """Save Excel with retry logic and backup."""
    backup_file = EXCEL_FILE.replace(".xlsx", "_backup.xlsx")

    for attempt in range(max_retries):
        try:
            temp_file = EXCEL_FILE.replace(".xlsx", "_temp.xlsx")
            df.to_excel(temp_file, index=False)

            if os.path.exists(EXCEL_FILE):
                try:
                    os.replace(EXCEL_FILE, backup_file)
                except PermissionError:
                    pass

            os.replace(temp_file, EXCEL_FILE)
            return True

        except PermissionError:
            if attempt < max_retries - 1:
                logging.warning(f"Save attempt {attempt + 1} failed - file may be open. Retrying in 2s...")
                time.sleep(2)
            else:
                logging.error(f"!!! CLOSE {EXCEL_FILE} in Excel so the script can save progress !!!")
                try:
                    df.to_excel(backup_file, index=False)
                except:
                    logging.error("Could not save backup either!")
                return False
    return False

# ---------------- CLAIM FLOW ---------------- #
def find_claim_card(page, reg_id):
    """Find claim card using multiple selector strategies."""
    selectors = [
        f"div:has-text('{reg_id}')",
        f"[class*='card']:has-text('{reg_id}')",
        f"[class*='claim']:has-text('{reg_id}')",
        f"tr:has-text('{reg_id}')",
    ]

    for selector in selectors:
        try:
            locator = page.locator(selector).first
            if locator.is_visible(timeout=2000):
                return locator
        except:
            continue

    return page.locator("div").filter(has_text=reg_id).first

def find_edit_button(claim_card):
    """Find edit/action button using multiple strategies."""
    selectors = [
        "svg path",
        "[id*='Path']",
        "button:has-text('Edit')",
        "[class*='edit']",
        "[class*='action']",
        "svg",
        "[role='button']",
    ]

    for selector in selectors:
        try:
            locator = claim_card.locator(selector).first
            if locator.is_visible(timeout=1000):
                return locator
        except:
            continue

    return claim_card.locator("[id*='Path']").first

def find_file_input(page):
    """Find file input using multiple strategies."""
    selectors = [
        "input#SupportingDoc2",
        "input[type='file'][id*='Supporting']",
        "input[type='file'][id*='Doc']",
        "input[type='file']",
    ]

    for selector in selectors:
        try:
            locator = page.locator(selector).first
            locator.wait_for(state="attached", timeout=5000)
            return locator
        except:
            continue

    raise Exception("Could not find file input element")

@retry_on_failure()
def process_claim(page, df, idx):
    row = df.loc[idx]
    reg_id = str(row["registeration_id"]).strip()
    doc_path = str(row["supporting_doc"]).strip()
    remarks = str(row["remarks"])

    try:
        logging.info(f"--- Processing Row {idx+1}: {reg_id} ---")

        if not doc_path or doc_path.lower() == "nan":
            raise FileNotFoundError(f"No document path specified for {reg_id}")
        if not os.path.exists(doc_path):
            raise FileNotFoundError(f"File not found: {doc_path}")

        # Search box using XPath
        search_box = page.locator(
            'xpath=//*[@id="root"]/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/input')
        search_box.click()
        search_box.fill("")  # Clear any existing text
        search_box.type(reg_id)
        logging.info(f"Entered registration ID: {reg_id}")

        # Click search button using XPath
        search_button = page.locator(
            'xpath=//*[@id="root"]/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/span')
        search_button.click()
        logging.info("Clicked search button")

        # Wait indefinitely for search results to load
        enter_claim_btn = page.locator('xpath=//*[@id="Path_98789"]')
        while not enter_claim_btn.is_visible():
            time.sleep(0.5)
        logging.info("Search results loaded - claim button visible")

        # Click the button to enter the claim
        wait_and_click(enter_claim_btn, "Enter Claim Button")

        # Wait for Other button to be visible
        other_btn = page.locator('xpath=//*[@id="root"]/div[2]/div/div[3]/div[4]/button')
        while not other_btn.is_visible():
            time.sleep(0.5)
        logging.info("Other button is visible")

        # Check if Other button is closed or open and click if closed
        other_btn_closed = page.locator('button.YruTh3yguplDeNOJ2EEp:has-text("Other")')
        other_btn_open = page.locator('button.V4XyRzf_IMRHlR8iun2f:has-text("Other")')

        if other_btn_closed.is_visible():
            wait_and_click(other_btn_closed, "Other Button (opening)")
            logging.info("Other button was closed - clicked to open")
        elif other_btn_open.is_visible():
            logging.info("Other button already open - moving forward")
        else:
            # Fallback: click the button anyway
            wait_and_click(other_btn, "Other Button")

        # Wait indefinitely for Query Response button to be visible
        query_response_btn = page.locator('xpath=//*[@id="root"]/div[2]/div/div[3]/div[4]/div/div/div/div/div/div/div/div/div/div[1]/h2/button')
        while not query_response_btn.is_visible():
            time.sleep(0.5)
        logging.info("Query Response button is visible")

        # Check if Query Response is closed or open
        query_response_closed = page.locator('button.accordion-button.collapsed:has-text("Query Response")')
        query_response_open = page.locator('button.accordion-button:not(.collapsed):has-text("Query Response")')

        if query_response_closed.is_visible():
            wait_and_click(query_response_closed, "Query Response Button (opening)")
            logging.info("Query Response was closed - clicked to open")
            # Wait for accordion content to be visible
            accordion_content = page.locator('xpath=//*[@id="root"]/div[2]/div/div[3]/div[4]/div/div/div/div/div/div/div/div/div/div[2]')
            while not accordion_content.is_visible():
                time.sleep(0.5)
            logging.info("Query Response accordion content is now visible")
        elif query_response_open.is_visible():
            logging.info("Query Response already open - moving forward")
        else:
            # Fallback: click the button anyway
            wait_and_click(query_response_btn, "Query Response Button")

        wait_for_network_idle(page)

        logging.info(f"Uploading file: {doc_path}")
        file_input = find_file_input(page)
        file_input.set_input_files(doc_path)

        wait_for_network_idle(page, timeout=10000)

        remarks_field = page.get_by_role("textbox", name="Remarks")
        if not remarks_field.is_visible():
            remarks_field = page.locator("textarea[name*='remark'], input[name*='remark']").first
        wait_and_fill(remarks_field, remarks, "Remarks Field")

        # Wait for SAVE button to be visible and click
        save_btn = page.locator('xpath=//*[@id="root"]/div[2]/div/div[3]/div[4]/div/div/div/div/div/div/div/div/div/div[2]/div/form/div/div[2]/button')
        while not save_btn.is_visible():
            time.sleep(0.5)
        logging.info("SAVE button is visible")
        wait_and_click(save_btn, "SAVE button")
        wait_for_network_idle(page)

        # Click HOME button to go back to claims list
        home_btn = page.locator('xpath=//*[@id="root"]/div[2]/div/div[1]/div[1]/div/div[1]/div/div[2]/svg')
        while not home_btn.is_visible():
            time.sleep(0.5)
        logging.info("HOME button is visible")
        wait_and_click(home_btn, "HOME button")
        wait_for_network_idle(page)

        # Navigate back to Claims Queried page for next iteration
        navigate_to_claims(page)
        logging.info("Navigated back to Claims Queried page")

        # TODO: Uncomment below after testing
        # wait_and_click(page.get_by_role("button", name="SUBMIT CLAIM"), "SUBMIT button")
        # wait_and_click(page.get_by_role("button", name="YES"), "Confirmation YES")
        # wait_for_network_idle(page)

        df.at[idx, "status"] = "SUCCESS"
        logging.info(f"DONE: Successfully processed {reg_id}")

    except Exception as e:
        error_msg = str(e)
        df.at[idx, "status"] = f"FAILED: {error_msg[:100]}"
        logging.error(f"ERROR on {reg_id}: {error_msg}")

        try:
            page.go_back()
            wait_for_network_idle(page)
        except:
            pass

    finally:
        save_excel(df)

# ---------------- SESSION HANDLING ---------------- #
def check_session_valid(page):
    """Check if the session is still valid."""
    try:
        if page.locator("text='Session Expired'").is_visible(timeout=1000):
            return False
        if page.locator("text='Please login'").is_visible(timeout=1000):
            return False
        if page.locator("input[name*='login'], input[name*='user']").is_visible(timeout=1000):
            return False
        return True
    except:
        return True

def navigate_to_claims(page):
    """Navigate to the claims queried page."""
    try:
        page.get_by_role("button", name="CLOSE").click(timeout=3000)
    except:
        pass

    wait_and_click(page.get_by_role("button", name="View More"), "View More")

    # Wait indefinitely until Claims Queried menu is visible
    claims_queried = page.get_by_text("Claims Queried")
    while not claims_queried.is_visible():
        time.sleep(0.5)
    logging.info("Claims Queried menu is now visible")

    wait_and_click(claims_queried, "Claims Queried Menu")
    wait_for_network_idle(page)

    try:
        page.locator("select").first.select_option("50")
    except:
        pass

# ---------------- MAIN FUNCTION ---------------- #
def run():
    if not os.path.exists(EXCEL_FILE):
        logging.error(f"File {EXCEL_FILE} not found!")
        return

    try:
        df = load_excel()
    except FileNotFoundError as e:
        logging.error(str(e))
        return
    except Exception as e:
        logging.error(f"Error loading Excel file: {e}")
        return

    # Start Playwright
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=HEADLESS)
    context = browser.new_context()
    page = context.new_page()
    page.set_default_timeout(TIMEOUT)

    try:
        page.goto("https://provider.nha.gov.in/", wait_until="domcontentloaded", timeout=120000)

        print("\n>>> STEP 1: LOG IN MANUALLY")
        input("Complete first captcha -> Press ENTER")

        login_field = page.get_by_role("textbox", name="Registered Mobile No/User ID")
        if not login_field.is_visible():
            login_field = page.locator("input[type='text']").first
        wait_and_fill(login_field, USERNAME, "Login ID")

        verify_btn = page.get_by_text("verify")
        if not verify_btn.is_visible():
            verify_btn = page.locator("button:has-text('Verify'), button:has-text('verify')").first
        wait_and_click(verify_btn, "Verify Button")

        input("Enter OTP & second captcha -> Press ENTER when logged in")

        navigate_to_claims(page)

        # Process Rows
        processed = 0
        failed = 0
        skipped = 0

        for idx in df.index:
            if str(df.at[idx, "status"]).upper() == "SUCCESS":
                skipped += 1
                continue

            if processed > 0 and processed % 5 == 0:
                if not check_session_valid(page):
                    logging.error("Session expired! Please restart the script.")
                    print("\n>>> SESSION EXPIRED - Please restart the script")
                    break

            process_claim(page, df, idx)

            if str(df.at[idx, "status"]).upper() == "SUCCESS":
                processed += 1
            else:
                failed += 1

        # Summary
        print(f"\n>>> PROCESSING COMPLETE")
        print(f"    Processed: {processed}")
        print(f"    Failed: {failed}")
        print(f"    Skipped (already done): {skipped}")
        print(f"    Remaining: {len(df) - processed - failed - skipped}")

    except Exception as e:
        logging.error(f"Error occurred: {e}")
        print(f"\n>>> ERROR: {e}")
        print("Browser will stay open for debugging")

    input("Press Enter to close browser...")
    context.close()
    browser.close()
    playwright.stop()

# ---------------- ENTRY POINT ---------------- #
if __name__ == "__main__":
    run()