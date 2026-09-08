import sys
import time
import os
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configure the logging system globally
logging.basicConfig(
    filename='app.log',         # Name of the log file
    filemode='a',              # 'a' to append logs, 'w' to overwrite each run
    format='%(asctime)s - %(levelname)s - %(message)s', # Log structure
    level=logging.INFO         # Capture INFO level messages and above
)

# Initialize driver
options = webdriver.ChromeOptions()
options.add_argument("--use-fake-ui-for-media-stream")
options.add_argument("--use-fake-device-for-media-stream")

driver = webdriver.Chrome(options=options)
driver.implicitly_wait(10)
driver.get("https://govthealth.cg.gov.in/uhsmis/#/auth")

# Define an explicit wait timeout
wait = WebDriverWait(driver, 10)

CHC = "CHC BISHRAMPUR"
NAME = "Heenam Kushwaha"
target_village = "Karampur"
target_date = "01-09-2026"

def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)

def select_angular_dropdown(placeholder_text, option_text):
    """
    Helper function to click an Angular Material dropdown by its placeholder text
    and select a specific option from the overlay panel.
    """
    dropdown_xpath = f"//mat-form-field[contains(., '{placeholder_text}')]//mat-select | //div[contains(text(), '{placeholder_text}')]"
    dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
    dropdown.click()
    
    option_xpath = f"//mat-option[contains(., '{option_text}')] | //span[contains(@class, 'mat-option-text') and contains(text(), '{option_text}')]"
    option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
    option.click()

def check_date_and_village_set():
    try:
        # 1. Locate the form fields safely using explicit visibility checks
        date_field = wait.until(EC.visibility_of_element_located((By.XPATH, "//mat-form-field[contains(., 'Select Planned Visit Date')]")))
        village_field = wait.until(EC.visibility_of_element_located((By.XPATH, "//mat-form-field[contains(., 'Select Village')]")))

        # 2. Extract class attributes to verify validation status
        date_classes = date_field.get_attribute("class")
        village_classes = village_field.get_attribute("class")

        # Check if both fields contain the Angular invalid class marker
        if "ng-invalid" in date_classes and "ng-invalid" in village_classes:
            logging.warning("Validation failure: both inputs highlighted red.")
            print("Both fields are highlighted in red (invalid status). Injecting target data...")
            
            # Wait for global Angular loader to disappear completely
            wait.until(EC.invisibility_of_element_located((By.TAG_NAME, "app-loader")))
            
            # Input Planned Visit Date via JavaScript execution
            date_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Visit Date')] | //mat-form-field[contains(., 'Visit Date')]//input")))
            driver.execute_script("arguments[0].value = arguments[1];", date_input, target_date)
            driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", date_input)
            print(f"Set planned visit date to: {target_date}")

            # Open Village Dropdown Panel
            village_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-select | //mat-form-field[contains(., 'Village')]//mat-select")))
            village_dropdown.click()

            # Handle Option Selection using case-insensitive validation framework
            option_xpath = f"//mat-option[contains(translate(., 'KARAMPUR', 'karampur'), '{target_village.lower()}')] | //mat-option//span[contains(translate(text(), 'KARAMPUR', 'karampur'), '{target_village.lower()}')]"
            option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
            option.click()
            print(f"Selected village: {target_village}")

            # Target the search/continue execution button element
            search_btn_xpath = (
                "//button[@type='submit' or contains(., 'Continue')]"
                " | //mat-form-field//following::button[contains(., 'Continue')]"
                " | //span[contains(text(), 'Continue')]/ancestor::button"
                " | //button[contains(@class, 'mat-focus-indicator') and contains(., 'Continue')]"
            )
            search_btn = wait.until(EC.presence_of_element_located((By.XPATH, search_btn_xpath)))
            search_btn.click()
            logging.info(f"Form corrected and resubmitted for date: {target_date}, village: {target_village}")
            return True
        else:
            print("One or both fields do not show a validation error layout.")
            return False
            
    except Exception as e:
        logging.error("Error executing verification checks or inputting alternative parameters", exc_info=True)
        return False

def login():
    try:
        logging.info("Starting automated state login sequence...")
        
        # Step 1: Select District
        select_angular_dropdown("Select District", "SURAJPUR  (सूरजपुर )")
        
        # Step 2: Select CHC/UPHC
        select_angular_dropdown("Select CHC/UPHC", CHC)
        
        # Step 3: Handle Connection Question
        select_angular_dropdown("Is SHC Directly Connected", "Yes")
        
        # Step 4: Select SHC/AAM
        select_angular_dropdown("Select SHC/AAM", "SHC KARAMPUR")

        # Select the Employee Radio Button Card
        employee_xpath = f"//*[contains(text(), '{NAME}')]"
        employee_label = wait.until(EC.element_to_be_clickable((By.XPATH, employee_xpath)))
        employee_label.click()
        print(f"Selected employee: {NAME}")

        # Clear standard structural browser elements or application banners
        try:
            close_button = driver.find_element(By.XPATH, "//button[@aria-label='Close installation prompt']")
            close_button.click()
        except:
            pass  # Bypass if installation modal prompt doesn't show up

        # Confirm Password Selection Tab
        password_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Login with Password')] | //div[contains(text(), 'Login with Password')]")))
        password_tab.click()

        # Fill Password Field
        password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password' or @placeholder='Password']")))
        password_field.clear()
        password_field.send_keys("Karampur@123")

        # Handle Captcha Input Target Location
        captcha_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Captcha')] | //input[@name='captcha']")))
        print("\n[PAUSE] Execution stopped automatically. Please input the graphic captcha string manually into the portal UI element, then press ENTER in this terminal shell...")
        
        # Wait for developer terminal key token submission before attempting login step execution
        input()
        
        # Execute authentication click actions
        login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Login') or @type='submit']")))
        login_btn.click()
        logging.info("Login action initialized.")

    except Exception as e:
        logging.critical("Fatal system failure occurred within the application login loop structure", exc_info=True)

# Main Execution Routine
if __name__ == "__main__":
    login()
    # If checking the form fields is part of post-login dashboard workflow:
    # mySleepFunction(5) # Wait for workspace elements to load
    # check_date_and_village_set()
