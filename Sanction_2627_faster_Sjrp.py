import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time, os, csv
from pathlib import Path
from PIL import Image
import pytesseract
import cv2
import requests
import threading
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.keys import Keys
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time



w_FY_type = ["2016-2017", "2017-2018","2018-2019", "2019-2020", "2020-2021", "2022-2023", "2024-2025", "2025-2026"] 
FY_type = ["2024-2025", "2025-2026", "PMAYG Cumulative progress till date" ]
w_Scheme_type = ["PMAYG"]
Scheme_type = ["PRADHAN MANTRI AWAAS YOJANA GRAMIN"]
State_type = ["CHHATTISGARH"]
District_type = ["SURAJPUR"]
block_type = ["BHAIYATHAN", "ODAGI", "PRATAPPUR", "PREMNAGAR", "RAMANUJNAGAR", "SURAJPUR"] #"BHAIYATHAN",
data_type = ["Sanctioned Year"]
Panchayat_type = ["All"]
Category_type = ["All"]
Progress_type = ["All"]
DOWNLOAD_DIR = r"C:\Users\HP\Downloads>" 
Service = Service(timeout = 300)
driver = webdriver.Chrome(service=Service)
driver.get("https://pmayg.dord.gov.in/netiay/masterlogin.aspx")
os.environ["OMP_THREAD_LIMIT"] = "1"
driver.page_load_strategy = 'eager'

WAIT_SECONDS = 5
USER_ID = "CH26"
PASSWORD = "District@2027"
FY = "2026-2027"    
driver.implicitly_wait(10) 
wait = WebDriverWait(driver, 15)

def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)

def solveLoginCaptcha():
    captcha_image = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_imgCaptcha")
    captcha_image.screenshot("captcha.png")    
    captcha_text = pytesseract.image_to_string("captcha.png", config='--psm 6', lang='eng')
    print("Captcha Text:", captcha_text.strip())

    captcha_input = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtCaptcha")
    # check if captcha_text is empty or not
    captcha_input.send_keys(captcha_text.strip())
    input("Press Enter after solving captcha and logging in...")  # Wait for user to solve captcha and log in
    return captcha_text  

def login():
    dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFinYear"))
    dropdown.select_by_visible_text(FY)

    username = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtUserName")
    username.send_keys(USER_ID)

    password = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtPassword")
    password.send_keys(PASSWORD)
    print("Solving Captcha...")
    print( solveLoginCaptcha())
    
    #button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
    #button.click()
    print("Logged in successfully!")
    # alert = driver.find_element(By.NAME, "btnClose")
    # if alert:
    #     alert.click()
    #     print("Alert closed.")


    #calling to download A2 Report1649	1540	1330
    #A2_Report()
    #Aawas_Plus_Report()
    #WorkinProgress_Report()
    Sanction_2627()

def A2_Report():
    print("Navigating to A2 Report...")
    mySleepFunction(WAIT_SECONDS)
    # Selcting A2 Report
    try: 
        link = driver.find_element(By.PARTIAL_LINK_TEXT, "High level physical progress report")
        link.click()
    except Exception as e:
        print("Error while navigating to A2 Report:", e)
        driver.quit()
        return
    finally:
        print("Navigation to A2 Report attempted.")

    print("A2 Report Selected")
    for fy in FY_type:
        dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFinYear"))
        dropdown.select_by_visible_text(fy)
        print(fy)
        
        for scheme in Scheme_type:
            dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlScheme"))
            dropdown.select_by_visible_text(scheme)
            print(scheme)
            
            for state in State_type:
                dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlState"))
                dropdown.select_by_visible_text(state)
                print(state)
                if(state == "CHHATTISGARH"):
                    #Get State Data
                    button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
                    button.click()
                    mySleepFunction(WAIT_SECONDS)
                    button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExport")
                    button.click()
                for district in District_type: 
                    dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlDistrict"))
                    dropdown.select_by_visible_text(district)
                    print(district)

                    button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
                    button.click()
                    button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExport") 
                    button.click()
                    mySleepFunction(WAIT_SECONDS)
                    
                    #download block data for
                    for block in block_type:
                        dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlBlock"))
                        dropdown.select_by_visible_text(block)
                        print(block)
                        button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
                        button.click()
                        mySleepFunction(WAIT_SECONDS)
                        button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExport") 
                        button.click()
                        mySleepFunction(WAIT_SECONDS)
                        print("Data Downloaded for " + fy + " " + scheme + " " + state + " " + district)
                         # Got back to Home page after downloading all data
    driver.find_element(By.LINK_TEXT, "Home").click()
    alert = driver.find_element(By.NAME, "btnClose")
    if alert:
        alert.click()
        print("Alert closed.")    



def get_valid_options(select_element):
    """Returns a list of (value, text) tuples excluding placeholder options."""
    select = Select(select_element)
    return [
        (opt.get_attribute("value"), opt.text.strip())
        for opt in select.options
        if opt.get_attribute("value") and not opt.text.strip().lower().startswith("select")
    ]


# def select_and_trigger_change(wait, locator, value):
#     """Selects an option by value and forces a JS 'change' event for AJAX listeners."""
#     el = wait.until(EC.element_to_be_clickable(locator))
#     Select(el).select_by_value(value)
#     driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", el)



def select_and_trigger_change(wait, locator, value):
    el = wait.until(EC.presence_of_element_located(locator))
    try:
        Select(el).select_by_value(value)
    except ElementNotInteractableException:
        driver.execute_script("""
            const sel = arguments;
            sel.value = arguments;
            sel.dispatchEvent(new Event('input', { bubbles: true }));
            sel.dispatchEvent(new Event('change', { bubbles: true }));
        """, el, value)










def wait_for_options_to_populate(wait, by, locator):
    """Waits until a dependent <select> element has more than 1 option loaded."""
    def _has_options(drv):
        try:
            el = drv.find_element(by, locator)
            return len(Select(el).options) > 1
        except StaleElementReferenceException:
            return False

    wait.until(_has_options)


def reopen_preferences_modal(wait):
    """Re-opens #preferencesModal in the current window/frame context."""
    driver.execute_script("""
        var modalEl = document.getElementById('preferencesModal');
        if (modalEl && typeof bootstrap !== 'undefined') {
            var modal = bootstrap.Modal.getOrCreateInstance(modalEl);
            modal.show();
        }
    """)
    wait.until(EC.visibility_of_element_located((By.ID, "preferencesModal")))

def find_beneficiary_context():
    """Searches all open browser tabs and iframes to locate #tbBeneficiaryBody tr."""
    # Check all open browser tabs/windows (newest tab first)
    for handle in reversed(driver.window_handles):
        driver.switch_to.window(handle)
        driver.switch_to.default_content()

        # 1. Check main document of this tab
        if driver.find_elements(By.CSS_SELECTOR, "#tbBeneficiaryBody tr"):
            return True

        # 2. Check top-level iframes in this tab
        for frame in driver.find_elements(By.TAG_NAME, "iframe"):
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)
            if driver.find_elements(By.CSS_SELECTOR, "#tbBeneficiaryBody tr"):
                return True

            # 3. Check nested iframes (iframe inside iframe)
            for sub_frame in driver.find_elements(By.TAG_NAME, "iframe"):
                driver.switch_to.frame(sub_frame)
                if driver.find_elements(By.CSS_SELECTOR, "#tbBeneficiaryBody tr"):
                    return True
                driver.switch_to.parent_frame()

    return False

CSV_FILE = "Sanction_2627_Stages_Report_Surajpur.csv"
csv_gp_file = "visited_Surajpur.csv"
STAGE_NAMES = [
    "Geo Tagged",
    "Job Card No. Mapped & Verified",
    "Aadhaar Seeding",
    "Aadhaar Matched With VBGRAMG",
    "Aadhaar Not Duplicate With PMAYG",
    "Aadhaar Not Duplicate With PMAY-U",
    "e-KYC Done",
    "AI Validation",
]


def save_row_to_csv(row_dict):
    """Appends a single beneficiary record to the CSV file immediately."""
    fieldnames = [
        "Block",
        "Gram Panchayat",
        "Scheme",
        "Social Category",
        "Reg ID",
        "Beneficiary Name",
        "Father's Name",
        "Village",
        "Overall Status",
        *STAGE_NAMES,
    ]
    file_exists = os.path.isfile(CSV_FILE)
    with open(CSV_FILE, mode="a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row_dict)


def extract_detail_modal_stages():
    """Extracts the 8 stage statuses from the currently visible View Detail modal."""
    # Wait until the 'Stages' section text renders inside the open modal
    end_time = time.time() + 10
    while time.time() < end_time:
        has_stages = driver.execute_script("""
            var modal = document.querySelector('.modal.show:not(#preferencesModal)');
            return modal && modal.innerText.includes('Geo Tagged');
        """)
        if has_stages:
            break
        time.sleep(0.3)

    # Extract Village + all Stage -> Yes/No pairs dynamically via JS
    extracted = driver.execute_script("""
        var modal = document.querySelector('.modal.show:not(#preferencesModal)');
        if (!modal) return {};

        var result = {};
        var lines = modal.innerText
            .split('\\n')
            .map(function(l) { return l.trim(); })
            .filter(function(l) { return l.length > 0; });

        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];

            // Capture Village value right after 'Village:' label
            if (line.startsWith('Village:') && i + 1 < lines.length) {
                result['Village'] = lines[i + 1];
            }

            // Check if the next line is 'Yes' or 'No' (with or without checkmark/cross icon text)
            if (i + 1 < lines.length) {
                var nextClean = lines[i + 1].replace(/[^a-zA-Z]/g, '').trim();
                if (nextClean === 'Yes' || nextClean === 'No') {
                    result[line] = nextClean;
                }
            }
        }
        return result;
    """)
    return extracted or {}

def click_confirm_sanction_if_active(driver, popup_timeout=5):
    try:
        btn = driver.find_element(By.XPATH, "//button[normalize-space()='Confirm Sanction Order']")
        if not btn.is_enabled() or "disabled" in (btn.get_attribute("class") or "").lower():
            print("Confirm Sanction Order button is disabled. Skipping.")
            return False
        else:
            driver.execute_script("arguments[0].click();", btn)
            print("Clicked 'Confirm Sanction Order'.")
            # Click 'Done' if it appears; otherwise press ESC to close any dummy/empty modal cleanly
            try:
                done_btn = WebDriverWait(driver, popup_timeout).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#sanctionModal #btnSanctionDone"))
                )
                driver.execute_script("arguments[0].click();", done_btn)
                print("  Clicked 'Done'.")
                return True
            except TimeoutException:
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                return False

    except (NoSuchElementException, TimeoutException):
        return False
      # Small pause to allow any modal to close

def process_beneficiary_card(wait, block_text, gp_text, scheme_text, cat_text):
    row_xpath = "//table[not(ancestor::div[contains(@class,'modal')])]//tbody/tr[td]"
    view_detail_xpath = ".//a[normalize-space()='View Detail'] | .//button[normalize-space()='View Detail']"

    # 1. Fast wait for badge OR rows (exits immediately once badge updates or rows appear)
    expected_total = None
    end_wait = time.time() + 5
    while time.time() < end_wait:
        try:
            badge_text = driver.execute_script(
                "const el = document.querySelector('div.target-badge'); return el ? el.innerText : '';"
            )
            match = re.search(r'/\s*(\d+)', badge_text)
            if match:
                expected_total = int(match.group(1))
                # If > 0, or if we already waited at least 1.2s and it's 0, stop waiting
                if expected_total > 0 or (time.time() > end_wait - 3.8):
                    break
        except Exception:
            pass
        time.sleep(0.2)

    print(f"  Location: {block_text} / {gp_text} | Expected total: {expected_total}")

    if expected_total == 0:
        has_rows = driver.execute_script(
            "return document.querySelectorAll('table:not(.modal table) tbody tr td').length > 0;"
        )
        if not has_rows:
            print(f"  [INFO] 0 records for {gp_text} | {cat_text}. Skipping.")
            return

    seen_reg_ids = set()
    stale_scroll_attempts = 0
    rows_to_save = []  # Batch CSV writes per category

    # 2. Fast batch scan + scroll loop
    while True:
        if expected_total is not None and len(seen_reg_ids) >= expected_total:
            break

        try:
            # Single DOM call per scroll batch instead of per-row re-queries
            visible_rows = driver.find_elements(By.XPATH, row_xpath)
        except StaleElementReferenceException:
            time.sleep(0.2)
            continue

        new_row_processed = False

        for row in visible_rows:
            try:
                # Extract all cells in one fast JS call (avoids 6 Selenium round-trips per row)
                cell_texts = driver.execute_script(
                    "return Array.from(arguments[0].querySelectorAll('td')).map(td => td.innerText.trim());",
                    row
                )
                reg_id = cell_texts[1] if len(cell_texts) > 1 else ""

                if not reg_id or reg_id in seen_reg_ids:
                    continue

                seen_reg_ids.add(reg_id)
                new_row_processed = True

                ben_name = cell_texts[2] if len(cell_texts) > 2 else ""
                father_name = cell_texts[3] if len(cell_texts) > 3 else ""
                status = cell_texts[5] if len(cell_texts) > 5 else ""

                row_links = row.find_elements(By.XPATH, view_detail_xpath)
                clicked = False

                if row_links:
                    print(f"[{len(seen_reg_ids)}/{expected_total or '?'}] View Detail -> {reg_id} | {ben_name}")
                    # Direct JS click works even without scrollIntoView
                    driver.execute_script("arguments[0].click();", row_links[0])
                    clicked = True
                else:
                    print(f"[{len(seen_reg_ids)}/{expected_total or '?'}] No Link -> {reg_id} | {status}")

                stages_data = extract_detail_modal_stages() if clicked else {}

                record = {
                    "Block": block_text,
                    "Gram Panchayat": gp_text,
                    "Scheme": scheme_text,
                    "Social Category": cat_text,
                    "Reg ID": reg_id,
                    "Beneficiary Name": ben_name,
                    "Father's Name": father_name,
                    "Village": stages_data.get("Village", ""),
                    "Overall Status": status,
                }
                for stage in STAGE_NAMES:
                    record[stage] = stages_data.get(stage, "")

                rows_to_save.append(record)

                if clicked:
                    print("      Stages:", {s: record[s] for s in STAGE_NAMES})
                    visible_modals = [
                        m for m in driver.find_elements(By.CSS_SELECTOR, ".modal.show:not(#preferencesModal)")
                        if m.is_displayed()
                    ]
                    if visible_modals:
                        active_modal = visible_modals[-1]
                        close_btns = active_modal.find_elements(
                            By.CSS_SELECTOR, "button[data-bs-dismiss='modal'], .btn-close"
                        )
                        if close_btns:
                            target_btn = close_btns[-1]
                            try:
                                target_btn.click()  # Native Selenium click (triggers Bootstrap listeners reliably)
                            except Exception:
                                driver.execute_script("arguments[0].click();", target_btn)
                        else:
                            # Fallback: force Bootstrap hide or ESC key
                            driver.execute_script("""
                                const m = arguments[0];
                                if (window.bootstrap && bootstrap.Modal) {
                                    const inst = bootstrap.Modal.getInstance(m) || new bootstrap.Modal(m);
                                    inst.hide();
                                }
                            """, active_modal)

                        # Wait until NO visible detail modals remain on screen
                        wait.until(lambda d: d.execute_script("""
                            return Array.from(document.querySelectorAll('.modal.show:not(#preferencesModal)'))
                                .filter(el => el.offsetParent !== null).length === 0;
                        """))
                        time.sleep(0.15)


            except StaleElementReferenceException:
                # DOM refreshed mid-batch; break to re-grab fresh visible_rows
                break

        # Flush batch to CSV immediately so progress is never lost
        if rows_to_save:
            for rec in rows_to_save:
                save_row_to_csv(rec)
            rows_to_save.clear()

        if expected_total is not None and len(seen_reg_ids) >= expected_total:
            break

        # 3. Fast scroll down using div.table-responsive or last row
        if not new_row_processed:
            stale_scroll_attempts += 1
            if stale_scroll_attempts >= 2:
                break
        else:
            stale_scroll_attempts = 0

        driver.execute_script("""
            const container = document.querySelector('div.table-responsive');
            if (container && container.scrollHeight > container.clientHeight) {
                container.scrollTop += container.clientHeight * 0.85;
            } else {
                const rows = document.querySelectorAll('table:not(.modal table) tbody tr');
                if (rows.length) rows[rows.length - 1].scrollIntoView({block: 'start'});
            }
        """)
        time.sleep(0.35)
    return click_confirm_sanction_if_active(driver)

def get_visited_gps(csv_path, current_block):
    """Reads all previously recorded GPs for the given block across all runs."""
    visited = set()
    if not os.path.exists(csv_path):
        return visited

    with open(csv_path, mode="r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            block_name = row.get("Block", "").strip().lower()
            gp_name = row.get("GP", "").strip().lower()
            gp_val = row.get("GP_Value", "").strip().lower()
            if block_name == current_block.strip().lower():
                if gp_name:
                    visited.add(gp_name)
                if gp_val:
                    visited.add(gp_val)
    return visited

def append_visited_gp(csv_path, current_block, gp_name, gp_val=""):
    """Appends a single completed GP to the CSV file immediately."""
    write_header = not os.path.exists(csv_path) or os.path.getsize(csv_path) == 0

    with open(csv_path, mode="a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["Block", "GP", "GP_Value"])
        if write_header:
            writer.writeheader()
        writer.writerow({
            "Block": current_block.strip(),
            "GP": gp_name.strip(),
            "GP_Value": str(gp_val).strip()
        })
        f.flush()  # Forces immediate disk write in case the script stops



def Sanction_2627():
    wait = WebDriverWait(driver, 15)

    print("Navigating to Sanction module...")
    #mySleepFunction(WAIT_SECONDS)
    # Selcting A2 Report
    try: 
        link = driver.find_element(By.PARTIAL_LINK_TEXT, "Sanction")
        link.click()
    except Exception as e:
        print("Error while navigating to Sanction module:", e)
        driver.quit()
        return
    finally:
        print("Navigation to Sanction module attempted.")

    print("Sanction module selected")


  # 1. Wait for #preferencesModal and Block dropdown to be visible
    wait.until(EC.visibility_of_element_located((By.ID, "preferencesModal")))
    block_locator = (By.ID, "blockSelect")
    block_el = wait.until(EC.visibility_of_element_located(block_locator))

    # Indexed XPaths for the 2nd, 3rd, and 4th <select> inside #preferencesForm
    gp_locator = (By.XPATH, "(//div[@id='preferencesForm']//select)[2]")
    scheme_locator = (By.XPATH, "(//div[@id='preferencesForm']//select)[3]")
    category_locator = (By.XPATH, "(//div[@id='preferencesForm']//select)[4]")
    proceed_btn_locator = (By.ID, "btnProceedPref")

    blocks = get_valid_options(block_el)
    print(f"Total Blocks found: {len(blocks)}")

    #[('3305009','SURAJPUR'),('3305010', 'PREMNAGAR'),('33050011','RAMANUJNAGAR'),('3305012', 'BHAIYATHAN'),('3305013', 'ODAGI'),('3305015', 'PRATAPPUR')
    # 2. Loop through all Blocks
    for block_val, block_text in  [('3305009', 'SURAJPUR')]:
        select_and_trigger_change(wait, block_locator, block_val)

        try:
            wait_for_options_to_populate(wait, *gp_locator)
        except TimeoutException:
            print(f"[SKIP] No Gram Panchayats loaded for Block: {block_text}")
            continue

        gps = get_valid_options(driver.find_element(*gp_locator))
        # --- Before starting the GP loop for the current block ---
        visited_gps = get_visited_gps(csv_gp_file, block_text)

        # 3. Loop through all Gram Panchayats
        for gp_val, gp_text in gps:
            gp_text_clean = gp_text.strip().lower()
            gp_val_clean = str(gp_val).strip().lower()

            # Skip if already recorded in any previous run
            if gp_text_clean in visited_gps or gp_val_clean in visited_gps:
                print(f"[SKIP] Already visited GP: {gp_text} (Block: {block_text})")
                continue
            append_visited_gp(csv_gp_file, block_text, gp_text, gp_val)  # Record this GP as visited

            select_and_trigger_change(wait, gp_locator, gp_val)

            try:
                wait_for_options_to_populate(wait, *scheme_locator)
            except TimeoutException:
                print(f"[SKIP] No Schemes loaded for GP: {gp_text}")
                continue

            schemes = get_valid_options(driver.find_element(*scheme_locator))

            # 4. Loop through all Schemes
            for scheme_val, scheme_text in schemes:
                select_and_trigger_change(wait, scheme_locator, scheme_val)

                try:
                    wait_for_options_to_populate(wait, *category_locator)
                except TimeoutException:
                    print(f"[SKIP] No Social Categories loaded for Scheme: {scheme_text}")
                    continue

                categories = get_valid_options(driver.find_element(*category_locator))

                # 5. Loop through all Social Categories
                for cat_val, cat_text in categories:
                     
                    select_and_trigger_change(wait, category_locator, cat_val)

                    # Wait for the new table rows to appear initially
                    time.sleep(1.5)

                    # ---> INSERT THE SCROLL BLOCK RIGHT HERE <---
                    # # Scroll down to force all lazy-loaded rows in the table to render in the DOM
                    last_height = driver.execute_script("return document.body.scrollHeight")
                    while True:
                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(0.8)
                        new_height = driver.execute_script("return document.body.scrollHeight")
                        if new_height == last_height:
                            break
                        last_height = new_height

                    # Scroll back to the top before counting and clicking
                    driver.execute_script("window.scrollTo(0, 0);")
                    time.sleep(0.5)
                    # table_container = driver.find_element(By.CSS_SELECTOR, "div.table-responsive")
                    # driver.execute_script("arguments.scrollTop += 400;", table_container)
                    # time.sleep(0.6)

                    # ---> END OF INSERTED BLOCK <---


                    print(f"\nApplying -> Block: {block_text} | GP: {gp_text} | Scheme: {scheme_text} | Category: {cat_text}")

                    # Click Proceed
                    proceed_btn = wait.until(EC.element_to_be_clickable(proceed_btn_locator))
                    proceed_btn.click()

                    # Accept the "Preferences Applied Successfully!" JS alert
                    alert = wait.until(EC.alert_is_present())
                    alert.accept()

                    # Wait for #preferencesModal to finish hiding
                    wait.until(EC.invisibility_of_element_located((By.ID, "preferencesModal")))

                    # 6. Extract rows and click all 'View Detail' links inside main.beneficiary-card
                    process_beneficiary_card(wait, block_text, gp_text, scheme_text, cat_text)
                    mySleepFunction(2)  # Small pause to ensure modal closes before next iteration

                    # 7. Re-open #preferencesModal for the next iteration
                    reopen_preferences_modal(wait)



def Aawas_Plus_Report():
    FY_type = ["2024-2025", "2025-2026", "PMAYG Cumulative progress till date" ]
    print("Navigating to Aawas+ Report...")
    mySleepFunction(WAIT_SECONDS)
    # Selcting A2 Report
    try: 
        link = driver.find_element(By.PARTIAL_LINK_TEXT, "AwaasPlus Physical Progress Report.")
        link.click()
    except Exception as e:
        print("Error while navigating to Aawas+ Report:", e)
        driver.quit()
        return
    finally:
        print("Navigation to Aawas+ Report attempted.")

    print("Aawas+ Report Selected")
    for fy in FY_type:
        dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFinYear"))
        dropdown.select_by_visible_text(fy)
        print(fy)
        button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
        button.click()
        button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExport") 
        button.click()
        mySleepFunction(WAIT_SECONDS)
    driver.find_element(By.LINK_TEXT, "Home").click()
    alert = driver.find_element(By.NAME, "btnClose")
    if alert:
        alert.click()
        print("Alert closed.")
    

# Login to the website

def wait_for_download_complete(download_dir, timeout=300):
    """
    Pauses execution until all active Chrome (.crdownload) files 
    disappear from the download directory.
    """
    print("Waiting for download to complete...")
    start_time = time.time()
    download_path = Path(download_dir)
    
    while True:
        # Look for any active Chrome temporary download files
        crdownload_files = list(download_path.glob("*.crdownload"))
        
        if not crdownload_files:
            # Look for any partially downloaded files (0 bytes) just to be absolutely sure
            tmp_files = [f for f in download_path.glob("WorkinProgress_*.xlsx") if f.stat().st_size == 0]
            if not tmp_files:
                print("Download complete!")
                return True
        
        # Prevent an infinite loop if the server crashes or stalls permanently
        if time.time() - start_time > timeout:
            print(f"Timed out waiting for download after {timeout} seconds.")
            return False
            
        time.sleep(1) # Check again every 1 second

def WorkinProgress_Report():
    #w_FY_type = ["2018-2019", "2019-2020", "2020-2021", "2022-2023", "2025-2026"] #"2016-2017", "2017-2018", 
    w_FY_type = ["2025-2026", "2024-2025"]
    #Manual intervention required for clicking download button due to timeout issues.
    #w_FY_type = ["2024-2025"]
    print("Navigating to Work in Progress Report...")
    mySleepFunction(WAIT_SECONDS)
    # Selcting A2 Report
    try: 
        link = driver.find_element(By.PARTIAL_LINK_TEXT, "Work Progress for PMAY-G")
        link.click()
    except Exception as e:
        print("Error while navigating to Work in Progress Report:", e)
        driver.quit()
        return
    finally:
        print("Navigation to Work in Progress Report attempted.")

    print("Work in Progress Report Selected")
    for fy in data_type:
        dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlyear"))
        dropdown.select_by_visible_text(fy)
        print(fy)
        
        for scheme in w_Scheme_type:
            dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rblScheme"))
            dropdown.select_by_visible_text(scheme)
            print(scheme)
            
            for fy in w_FY_type:
                dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFinYear"))
                dropdown.select_by_visible_text(fy)
                print(fy)
                for block in block_type:
                    dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlBlock"))
                    dropdown.select_by_visible_text(block)
                    print(block)
                    for panchayat in Panchayat_type:
                        dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlPanch"))
                        dropdown.select_by_visible_text(panchayat)
                        print(panchayat)
                        for category_type in Category_type:
                            dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlCat"))
                            dropdown.select_by_visible_text(category_type)
                            print(category_type)
                            for progress_type in Progress_type:
                                dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlSanction"))
                                dropdown.select_by_visible_text(progress_type)
                                print(progress_type)
                                #download block data for
                                button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExportexcel")
                                button.click()
                                mySleepFunction(5)
                                new_file_name = f"WorkinProgress_{block}_{fy}"
                                rename_latest_download(new_file_name)  
                                mySleepFunction(2)                             

                                # # NEW: Dynamically wait instead of using a static sleep timer
                                # if wait_for_download_complete(DOWNLOAD_DIR, timeout=15):
                                #     new_file_name = f"WorkinProgress_{block}_{fy}"
                                #     rename_latest_download(new_file_name)
                                # else:
                                #     print(f"Skipping rename for {block} due to download timeout.")
                                print("Data Downloaded for " + fy + " " + scheme + " " + fy + " " + block + " " + panchayat + " " + category_type + " " + progress_type)


def rename_latest_download(new_filename):
    # 1. Automatically locate the default user Downloads folder
    download_dir = Path.home() / "Downloads"
    move_dir = Path.home() / "Downloads/WIP"  # Adjust this path if your downloads go to a different folder
    
    # 2. Gather all items in the directory that are files
    files = [f for f in download_dir.iterdir() if f.is_file()]
    
    if not files:
        print("No files found in the Downloads folder.")
        return

    # 3. Find the most recently modified file
    latest_file = max(files, key=os.path.getmtime)
    
    # Keep the original file extension (e.g., .pdf, .zip, .csv)
    file_extension = latest_file.suffix
    
    # 4. Construct the new path
    new_file_path = move_dir / f"{new_filename}{file_extension}"
    
    # 5. Rename the file safely
    try:
        if not new_file_path.exists():
            latest_file.rename(new_file_path)
            print(f"Successfully renamed:\nFrom: '{latest_file.name}'\nTo:   '{new_file_path.name}'")
        else:
            print(f"Error: A file named '{new_file_path.name}' already exists.")
    except Exception as e:
        print(f"An error occurred: {e}")



login()
mySleepFunction(WAIT_SECONDS)
driver.quit()
