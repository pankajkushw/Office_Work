import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time, os



# Initialize driver
driver = webdriver.Chrome()
driver.get("https://govthealth.cg.gov.in/uhsmis/#/auth")

# Define an explicit wait timeout
wait = WebDriverWait(driver, 10)


CHC = "CHC BISHRAMPUR"
NAME = "Heenam Kushwaha"
driver.implicitly_wait(10) 
def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)# Define an explicit wait timeout
wait = WebDriverWait(driver, 10)

def select_angular_dropdown(placeholder_text, option_text):
    """
    Helper function to click an Angular Material dropdown by its placeholder text
    and select a specific option from the overlay panel.
    """
    # 1. Locate and click the dropdown trigger container based on its placeholder label text
    dropdown_xpath = f"//mat-form-field[contains(., '{placeholder_text}')]//mat-select | //div[contains(text(), '{placeholder_text}')]"
    dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
    dropdown.click()
    
    # 2. Wait for the material option panel overlay to pop up and click the matching choice
    option_xpath = f"//mat-option[contains(., '{option_text}')] | //span[contains(@class, 'mat-option-text') and contains(text(), '{option_text}')]"
    option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
    option.click()

def login():
    try:
        # Step 1: Select District (Replace 'Raipur' with your actual district value)
        select_angular_dropdown("Select District", "SURAJPUR  (सूरजपुर )")
        
        # Step 2: Select CHC/UPHC
        select_angular_dropdown("Select CHC/UPHC", CHC)
        
        # Step 3: Handle "Is SHC Directly Connected to CHC/UPHC?"
        select_angular_dropdown("Is SHC Directly Connected", "Yes")
        
        # Step 4: Select SHC/AAM
        select_angular_dropdown("Select SHC/AAM", "SHC KARAMPUR")

        # 1. Select the Employee Radio Button Card
        # Locates the container card by the specific employee's name text
        employee_name = "Heenam Kushwaha"  # Change to the target employee
        employee_xpath = f"//*[contains(text(), '{employee_name}')]"
        
        employee_label = wait.until(EC.element_to_be_clickable((By.XPATH, employee_xpath)))
        employee_label.click()
        print(f"Selected employee: {employee_name}")

        # 2. Confirm "Login with Password" Tab is Selected
        # Clicks the tab option if it isn't set by default
        password_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Login with Password')] | //div[contains(text(), 'Login with Password')]")))
        password_tab.click()

        # 3. Fill Password Field
        password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password' or @placeholder='Password']")))
        password_field.clear()
        password_field.send_keys("Karampur@123")

        # 4. Handle Captcha Input Field
        # Locates the input element associated with the visual label "Captcha"
        captcha_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Captcha')] | //input[@name='captcha']")))

        # Note: Automated captcha breaking requires an OCR service. 
        # For testing, you can pause execution here to type it manually:
        captcha_code = input("Please look at the browser window and type the displayed Captcha code: ")
        captcha_field.send_keys(captcha_code)

        # 5. Click the Final Login Button
        login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Login')] | //span[contains(text(), 'Login')]/ancestor::button")))
        login_btn.click()
        print("Login submission executed.")

        # Login Completed, Select Program

        # Waits up to 10 seconds for the "Go to Leprosy Abhiyan" button to become clickable
        leprosy_program_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[contains(., 'Go to Leprosy Abhiyan')] | //a[contains(., 'Go to Leprosy Abhiyan')] | //*[text()='Go to Leprosy Abhiyan']"
        )))


        # Clicks the button to open the dashboard module
        leprosy_program_btn.click()
        print("Successfully navigated to the Leprosy Abhiyan module.")


        mySleepFunction(5)
        # Locates the specific calendar grid cell for the 1st
        # by isolating elements that contain the literal text block "1"
        date_xpath = (
            "//div[contains(@class, 'calendar')]//*[text()='1'] | "
            "//span[text()='1'] | "
            "//*[normalize-space(text())='1']"
        )
        
        # Wait until the cell element is visible and ready to be tapped
        date_element = wait.until(EC.element_to_be_clickable((By.XPATH, date_xpath)))
        date_element.click()
        print("Successfully selected September 1st.")

        # Locates the "Go to Entry Page" button using text-based matching
        entry_page_xpath = (
            "//button[contains(., 'Go to Entry Page')] | "
            "//a[contains(., 'Go to Entry Page')] | "
            "//*[text()='Go to Entry Page']"
        )
        
        # Wait up to 10 seconds for the button to be visible and clickable
        entry_page_btn = wait.until(EC.element_to_be_clickable((By.XPATH, entry_page_xpath)))
        entry_page_btn.click()
        print("Successfully clicked 'Go to Entry Page' button.")


        # Replace with the name of the village you want to choose from the modal dropdown options
        target_village = "Karampur" 

        # 1. Locate and click the Angular dropdown menu container
        dropdown_xpath = "//mat-select[contains(., 'Select Village')] | //div[contains(text(), 'Select Village')] | //mat-form-field[contains(., 'Select Village')]"
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
        dropdown.click()
        print("Dropdown opened.")

        # 2. Wait for the overlay option panel to pop up and click the targeted choice
        option_xpath = f"//mat-option[contains(., '{target_village}')] | //span[contains(@class, 'mat-option-text') and contains(text(), '{target_village}')]"
        option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
        option.click()
        print(f"Selected village: {target_village}")
        
        # 3. Locate and click the "Continue" button
        # Using text matching helps isolate it even if it changes from disabled to active state
        continue_btn_xpath = "//button[contains(., 'Continue')] | //span[contains(text(), 'Continue')]/ancestor::button"
        continue_btn = wait.until(EC.element_to_be_clickable((By.XPATH, continue_btn_xpath)))
        continue_btn.click()
        print("Successfully clicked 'Continue'.")

        # Replace this list with your actual target Ration Card numbers
        ration_cards = ["102345678912", "102345678913", "102345678914"]

        for card_number in ration_cards:
            print(f"Executing sequence for card entry: {card_number}")
            try:
                # Targets the input field directly associated with the ID card icon,
                # explicitly avoiding any date picker fields containing calendar icons.
                input_xpath = (
                    "//input[@type='text' and not(ancestor::mat-form-field[.//mat-datepicker-toggle]) and not(contains(@placeholder, 'Date'))]"
                    " | //mat-label[contains(., 'Card') or contains(., 'ABHA')]/ancestor::mat-form-field//input"
                    " | (//mat-form-field//input)[last()]"
                )
                
                # Wait until the true search input field is present
                search_field = wait.until(EC.presence_of_element_located((By.XPATH, input_xpath)))
                
                # Inject values directly using JavaScript to prevent calendar overlays from popping up
                # Inject values directly into the input using the proper indexed arguments
                driver.execute_script("arguments[0].value = arguments[1];", search_field, card_number)
                driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", search_field)
                driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", search_field)

                print("Card number successfully injected into the correct Search box.")
                
                # 1. Broadly target the search button card using text and style classes
                search_btn_xpath = (
                    "//button[@type='submit' or contains(., 'Search')]"
                    " | //mat-form-field//following::button[contains(., 'Search')]"
                    " | //span[contains(text(), 'Search')]/ancestor::button"
                    " | //button[contains(@class, 'mat-focus-indicator') and contains(., 'Search')]"
                )
                
                # 2. Wait until the button is present in the DOM layout
                search_btn = wait.until(EC.presence_of_element_located((By.XPATH, search_btn_xpath)))
                
                # 3. Clean click execution strategy
                try:
                    # Try a standard driver click first to allow Angular event bubbles to fire naturally
                    search_btn.click()
                    print("Standard browser search button click executed.")
                #Screening Logic will go here$$$$$$$$$$$$$$$$$$$$$$$$$$$$














                #Screening Logic will go here$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                except Exception:
                    # Fallback: If intercepted, clear focus variables and force injection click directly on the button node
                    print("Standard click blocked. Executing native node JavaScript click...")
                    driver.execute_script("arguments[0].focus();", search_btn)
                    driver.execute_script("arguments[0].click();", search_btn)
                    print("Search button JavaScript click forced.")
                
                time.sleep(3)
                
            except Exception as e:
                import traceback
                print(f"Pipeline crashed for card: {card_number}")
                print(traceback.format_exc())





    except Exception as e:
        print(f"An error occurred while filling the form: {e}")




# start executing
login()