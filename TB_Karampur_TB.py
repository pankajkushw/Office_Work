import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
import time, os, logging, sys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import base64



# Initialize driver
options = webdriver.ChromeOptions()
options.add_argument("--use-fake-ui-for-media-stream")
options.add_argument("--use-fake-device-for-media-stream")



driver = webdriver.Chrome(options=options)
driver.get("https://govthealth.cg.gov.in/uhsmis/#/auth")

# Define an explicit wait timeout
wait = WebDriverWait(driver, 10)


CHC = "CHC BISHRAMPUR"
NAME = "Heenam Kushwaha"
target_village = "Balrampurr"
target_date = "08-oct-2026" 
driver.implicitly_wait(10) 
def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)# Define an explicit wait timeout
wait = WebDriverWait(driver, 10)

# Configure the logging system
logging.basicConfig(
    filename='Birpur08102026.log',         # Name of the log file
    filemode='a',              # 'a' to append logs, 'w' to overwrite each run
    format='%(asctime)s - %(levelname)s - %(message)s', # Log structure
    level=logging.INFO         # Capture INFO level messages and above
)


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

def check_date_and_village_set():
    # Define target values
    # 1. Locate the form fields
    # Date container element mapping to the mat-form-field
    date_field = driver.find_element(By.XPATH, "//mat-form-field[contains(., 'Select Planned Visit Date')]")
    # Village container element mapping to the mat-form-field
    village_field = driver.find_element(By.XPATH, "//mat-form-field[contains(., 'Select Village')]")

    # 2. Extract class attributes to verify the "ng-invalid" validation status (red highlight indicator)
    date_classes = date_field.get_attribute("class")
    village_classes = village_field.get_attribute("class")

    if "ng-invalid" in date_classes and "ng-invalid" in village_classes:

        # 1. Input Target Date via JavaScript (Ensures Angular detects the change)
        date_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "mat-input-2"))
        )
        driver.execute_script(
            """
            var element = arguments[0];
            var val = arguments[1];
            element.value = val;
            element.dispatchEvent(new Event('input', { bubbles: true }));
            element.dispatchEvent(new Event('change', { bubbles: true }));
            """, 
            date_input, 
            target_date
        )

        # 2. Select Village from the Angular Material Dropdown
        # First, click the dropdown trigger container to open the option panel
        village_dropdown = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "mat-select[formcontrolname='village']"))
        )
        village_dropdown.click()

        # Wait for the option panel overlay to appear and click the target village option
        # Using dynamic text mapping to locate the exact village row
        village_option = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, f"//mat-option//span[contains(text(), '{target_village}')]"))
        )
        village_option.click()

        # 3. Click the Continue Button
        # Using the specific text inside the button container
        continue_button = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continue')]"))
        )
        continue_button.click()

    else:
        print(f"One or both fields do not show a validation error layout.(Line: {sys._getframe().f_lineno})")
        return False



def for_single():
    select_angular_dropdown("Belongs to PVTG Category", "No")
    time.sleep(0.5)
    
    # Wait for the option list overlay to appear and click "No"
    select_angular_dropdown("Suspected", "No")
    
    # Small pause to let the overlay close cleanly
    time.sleep(0.5)

    #select_angular_dropdown("गर्भवती / स्तनपान कराने वाली", "No")
    try:
        # 1. 3 सेकंड का एक छोटा वेट लगाएँ ताकि चेक किया जा सके कि dropdown स्क्रीन पर मौजूद है या नहीं
        dropdown_label = "गर्भवती / स्तनपान कराने वाली"
        select_xpath = f"//mat-select[contains(normalize-space(.), '{dropdown_label}')] | //div[contains(normalize-space(.), '{dropdown_label}')]//mat-select"
        
        # यदि element दिखाई देता है, तो इसे variable में स्टोर करें
        is_visible = WebDriverWait(driver, 5).until( #from 3 to 10
            EC.presence_of_element_located((By.XPATH, select_xpath))
        )
        
        # 2. केवल element मिलने पर ही फ़ंक्शन को कॉल करें
        select_angular_dropdown("गर्भवती / स्तनपान कराने वाली", "No")
        print("गर्भवती / स्तनपान कराने वाली dropdown सफलतापूर्वक सेट कर दिया गया है।")

    except Exception:
        # अगर element 3 सेकंड में नहीं मिलता, तो script बिना क्रैश हुए इसे छोड़ देगी
        print("गर्भवती / स्तनपान कराने वाली dropdown स्क्रीन पर दिखाई नहीं दिया। आगे बढ़ रहे हैं...")

    
    time.sleep(0.5)

    # ----------------------------------------------------
    # 3. Tick the Consent Checkbox
    # ----------------------------------------------------
    # Target the inner invisible input or the mat-checkbox label component
    # 1. Locate the checkbox element safely using dynamic layout variations
    checkbox = wait.until(EC.presence_of_element_located((
        By.XPATH, "//mat-checkbox//input[@type='checkbox'] | //mat-checkbox | //input[@type='checkbox']"
    )))

    # 2. Scroll to the element to make sure it is in view
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)

    # 3. Force the click event via JavaScript to avoid element click intercepted errors
    driver.execute_script("arguments[0].click();", checkbox)

    # ----------------------------------------------------
    # 4. Upload a Dummy Picture
    # ----------------------------------------------------
    # --- 1. File Upload Phase ---
    dummy_image_path = os.path.abspath("temp_placeholder.jpg")
    if not os.path.exists(dummy_image_path):
        with open(dummy_image_path, "wb") as f:
            f.write(b"\xFF\xD8\xFF\xE0\x00\x10JFIF\x00\x01\x01\x01\x00`\x00`\x00\x00\xFF\xDB\x00C\x00\x08\x06\x06\x07\x06\x05\x08\x07\x07\x07\t\t\x08\n\x0C\x14\r\x0C\x0B\x0B\x0C\x19\x12\x13\x0F\x14\x1D\x1A\x1F\x1E\x1D\x1A\x1C\x1C $.' \",#\x1C\x1C(7),01444\x1F'9=82<.342\xFF\xC0\x00\x0B\x08\x00\x01\x00\x01\x01\x01\x11\x01\xFF\xC4\x00\x15\x00\x01\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\xFF\xDA\x00\x0C\x01\x01\x00\x00\x3F\x00\xB2\xC0\xFF\xD9")

    # Target the hidden browser file channel input directly
    photo_input = driver.find_element(By.XPATH, "//input[@type='file']")
    photo_input.send_keys(dummy_image_path)
    print("Placeholder photo successfully routed to file stream.")

    # --- 2. Corrected Synchronization Wait ---
    # Use the global 'wait' object instance to keep timeout metrics uniform.
    # This explicitly waits for the image rendering preview frame to appear in the container layout.
    wait.until(EC.presence_of_element_located((
        By.XPATH, "//div[contains(@class, 'image')]//img | //img[not(@id) and @src] | //*[contains(@class, 'preview')]"
    )))
    print("Form validation refreshed: Photo preview detected.")

    # --- 3. Click Execution ---
    submit_btn = wait.until(EC.presence_of_element_located((
        By.XPATH, "//button[contains(normalize-space(.), 'Submit Leprosy Report')]"
    )))
    driver.execute_script("arguments[0].click();", submit_btn)
    print("Form submission executed successfully.")


    # 1. Explicitly wait until the SweetAlert confirm button is interactive on the screen viewport
    yes_save_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
    
    # 2. Execute the click using JavaScript to guarantee execution through the backdrop fade overlay
    driver.execute_script("arguments[0].click();", yes_save_btn)
    print("Confirmation modal 'Yes, Save' button successfully clicked.")


    try:
        # 1. Target the button via its unique SweetAlert confirmation class
        success_ok_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
        
        # 2. Fix: Corrected syntax using arguments[0] to run the native browser click track
        driver.execute_script("arguments[0].click();", success_ok_btn)
        print("Success popup 'OK' button clicked via corrected JS call.")
        
    except Exception:
        # Fallback Option: If the overlay layer blocks it, move the real pointer directly to the center and click
        print("JavaScript click fallback initiated...")
        success_ok_btn = driver.find_element(By.CSS_SELECTOR, "button.swal2-confirm")
        ActionChains(driver).move_to_element(success_ok_btn).click().perform()
        print("Success popup 'OK' button forcefully clicked via Actions API.")    


def login():
    # Check if the root logger has any handlers configured
    if not logging.root.handlers:
    # Call basicConfig to add a console handler with a pre-defined format
        logging.basicConfig(
            format='%(asctime)s - %(levelname)s - %(message)s',
            level=logging.INFO
        )

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

        # To click the pink/red "X" close button:
        close_button = driver.find_element(By.XPATH, "//button[@aria-label='Close installation prompt']")
        close_button.click()

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

        # Locate the precise button using the exact text inside the inner tag
        login_btn = driver.find_element(By.XPATH, "//button[span[text()='Login']]")

        # Force execution bypassing DOM layout restrictions
        driver.execute_script("arguments[0].click();", login_btn)
        print("Login button forcefully clicked via JS.")

        # Login Completed, Select Program

        leprosy_program_btn = wait.until(EC.presence_of_element_located((
            By.XPATH, "//*[normalize-space(.)='Go to TB']"
        )))

        # Force the click via JavaScript to completely bypass any Angular Material 
        # click interception layers or animation delays.
        driver.execute_script("arguments[0].click();", leprosy_program_btn)


        #mySleepFunction(5)
        # Locates the specific calendar grid cell for the 8th
        # by isolating elements that contain the literal text block "8"
        date_xpath = (
            "//div[contains(@class, 'calendar')]//*[text()='8'] | "
            "//span[text()='8'] | "
            "//*[normalize-space(text())='8']"
        )
        mySleepFunction(1)
        # Wait until the cell element is visible and ready to be tapped
        date_element = wait.until(EC.element_to_be_clickable((By.XPATH, date_xpath)))
        date_element.click()
        print("Successfully selected September 8th.")

        # Locates the "Go to Entry Page" button using text-based matching
        entry_page_xpath = (
            "//button[contains(., 'Go to Entry Page')] | "
            "//a[contains(., 'Go to Entry Page')] | "
            "//*[text()='Go to Entry Page']"
        )
        
        # Go to entry page
        entry_page_btn = wait.until(EC.element_to_be_clickable((By.XPATH, entry_page_xpath)))
        entry_page_btn.click()
        print(f"Successfully clicked 'Go to Entry Page' button.line: {sys._getframe().f_lineno}")

        # continue page
        continue_btn_xpath = "//button[contains(., 'Continue')] | //span[contains(text(), 'Continue')]/ancestor::button"
        continue_btn = wait.until(EC.element_to_be_clickable((By.XPATH, continue_btn_xpath)))
        continue_btn.click()
        print(f"Successfully clicked 'Continue'. (line: {sys._getframe().f_lineno})")

        # Replace this list with your actual target Ration Card numbers
        ration_cards = ['226489587274', '226489665424', '226489700178', '226489709765', '226489813931', '226489913534', '226484565907']
        
        for card_number in ration_cards:
            
            print(f"Executing sequence for card entry: {card_number}")
            logging.info(f"Opening:  {card_number}")
            try:
                print(f"Date and Village fields were not set correctly. Re-injecting values for card: {card_number} (line: {sys._getframe().f_lineno})")
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

                print(f"Card number successfully injected into the correct Search box. (line: {sys._getframe().f_lineno})")
                search_btn_xpath = (
                    "//button[@type='submit' or contains(., 'Search')]"
                    " | //mat-form-field//following::button[contains(., 'Search')]"
                    " | //span[contains(text(), 'Search')]/ancestor::button"
                    " | //button[contains(@class, 'mat-focus-indicator') and contains(., 'Search')]"
                )
                search_btn = wait.until(EC.presence_of_element_located((By.XPATH, search_btn_xpath)))
                search_btn.click()
                print(f"Search Member of Ration Card. (line: {sys._getframe().f_lineno})")


                #Screening Logic will go here$$$$$$$$$$$$$$$$$$$$$$$$$$$
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "custom-table")))
                # Target the buttons precisely using the class name 'action-btn' shown in your HTML
                select_buttons = driver.find_elements(By.CSS_SELECTOR, "button.action-btn")
                total_buttons = len(select_buttons)

                print(f"Found {total_buttons} matching 'Select' buttons. line: {sys._getframe().f_lineno}")
                logging.info(f"Member in :  {card_number} : {total_buttons}")

                counter = 0
                
                for i in range(total_buttons):
                    print(f"Processing Select button #{i + 1} of {total_buttons} for card {card_number}. line: {sys._getframe().f_lineno}")
                    
                    try:
                        #Second time date of visit, village has to be set and search again
                        if counter > 0:


                            search_btn_xpath = (
                                "//button[@type='submit' or contains(., 'Search')]"
                                " | //mat-form-field//following::button[contains(., 'Search')]"
                                " | //span[contains(text(), 'Search')]/ancestor::button"
                                " | //button[contains(@class, 'mat-focus-indicator') and contains(., 'Search')]"
                            )
                            # 2. Wait until the button is present in the DOM layout
                            search_btn = wait.until(EC.presence_of_element_located((By.XPATH, search_btn_xpath)))
                            # Try a standard driver click first to allow Angular event bubbles to fire naturally
                            search_btn.click()

                        # Re-fetch the elements inside the loop to ensure they are fresh
                        buttons = driver.find_elements(By.CSS_SELECTOR, "button.action-btn")
                        current_button = buttons[i]
                        
                        # Scroll the specific button into view before interacting
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", current_button)
                        time.sleep(0.5)
                        
                        print(f"Clicking Select button #{i + 1}... (line: {sys._getframe().f_lineno})")
                        
                        # Use a reliable JavaScript click to bypass overlapping Angular components or overlay layers
                        driver.execute_script("arguments[0].click();", current_button)

  
                         # --- REST OF FORM SUBMISSION ROUTINE CONTINUES BELOW ---
                        print(f"Proceeding with normal report creation actions... (line: {sys._getframe().f_lineno})")

                        try:
                            try:
                            # 1. Locate the ABHA ID field using an exception-safe layout locator
                                # Tries the expected dynamic ID first, falls back to formcontrolname if the ID shifts
                                abha_input = WebDriverWait(driver, 1).until(
                                    EC.visibility_of_element_located((
                                        By.XPATH, "//input[@formcontrolname='abhaId' or @id='mat-input-58']"
                                    ))
                                )

                                # 2. Extract current value cleanly
                                current_value = abha_input.get_attribute("value")
                                current_value = current_value.strip() if current_value else ""

                                # 3. Process if blank
                                if not current_value:
                                    random_abha_id = "".join([str(rand_mod.randint(0, 9)) for _ in range(14)])
                                    
                                    # Safe positional array arguments injection
                                    driver.execute_script(
                                        """
                                        var element = arguments[0];
                                        var val = arguments[1];
                                        if (element) {
                                            element.value = val;
                                            element.dispatchEvent(new Event('input', { bubbles: true }));
                                            element.dispatchEvent(new Event('change', { bubbles: true }));
                                        }
                                        """, 
                                        abha_input, 
                                        random_abha_id
                                    )
                                    print(f"ABHA ID was blank. Filled with random value: {random_abha_id}")
                                else:
                                    print(f"ABHA ID already has a value: {current_value}")

                            except Exception as e:
                                # Captures timeouts or stale element references without stopping the pipeline execution
                                print(f"Bypassed ABHA ID block checking error: {e}")
                                logging.warning(f"Could not check or update ABHA ID field layout: {e}")  


                            print(f"screening for {card_number}: {i}")
                            # ----------------------------------------------------
                            # 1. Select "No" in the "Belongs to PVTG Category" Dropdown
                            # ----------------------------------------------------
                            # Locate and click the PVTG dropdown container to expand the options
                            select_angular_dropdown("Belongs to PVTG Category", "No")
                            time.sleep(0.5)

                            # 1. "जोखिम मापदंड नहीं है" चेकबॉक्स को ढूंढें और क्लिक करें
                            jokhim_checkbox = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//mat-checkbox[contains(., 'जोखिम मापदंड नहीं है')]//label"))
                            )
                            jokhim_checkbox.click()

                            # 2. "कोई लक्षण नहीं" चेकबॉक्स को ढूंढें और क्लिक करें
                            lakshan_checkbox = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//mat-checkbox[contains(., 'कोई लक्षण नहीं')]//label"))
                            )
                            lakshan_checkbox.click()
                            time.sleep(0.5)

                            # 1. Prepare your local file path
                            abs_file_path = os.path.abspath("600-kb.jpg")  # Ensure this file exists in your workspace

                            # 2. Convert the image to a base64 Data URL
                            with open(abs_file_path, "rb") as image_file:
                                encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
                                base64_data_url = f"data:image/jpeg;base64,{encoded_string}"

                            # 3. Handle standard hidden input file upload if present
                            file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
                            if file_inputs:
                                file_inputs[0].send_keys(abs_file_path)

                            # 4. Inject the base64 data directly into the preview element and notify Angular
                            js_injection_script = """
                                var uploadBtn = document.querySelector(".upload_btn");
                                if (uploadBtn) {
                                    var container = uploadBtn.closest('div, .row');
                                    var previewImg = container ? container.querySelector("img") : null;
                                    
                                    if (previewImg) {
                                        previewImg.src = arguments[0];
                                        
                                        // Dispatch multiple events to update Angular framework model state
                                        var events = ['change', 'input', 'blur'];
                                        events.forEach(function(name) {
                                            var ev = new Event(name, { bubbles: true, cancelable: true });
                                            previewImg.dispatchEvent(ev);
                                        });
                                        console.log("Visual preview successfully updated.");
                                    } else {
                                        console.log("No <img> tag found. Data sent to structural elements.");
                                    }
                                } else {
                                    console.error("Could not locate '.upload_btn' element.");
                                }
                            """

                            driver.execute_script(js_injection_script, base64_data_url)

                            # ----------------------------------------------------
                            # 3. Tick the Consent Checkbox
                            # ----------------------------------------------------
                            # Target the inner invisible input or the mat-checkbox label component
                            # 1. Locate the checkbox element safely using dynamic layout variations
                            # 2. "मैं अपनी सहमति से अपनी फोटो खिंचवाने की अनुमति देता/देती हूँ" चेकबॉक्स को ढूंढें और क्लिक करें
                            lakshan_checkbox = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//mat-checkbox[contains(., 'मैं अपनी सहमति से अपनी फोटो खिंचवाने की अनुमति देता/देती हूँ')]//label"))
                            )
                            lakshan_checkbox.click()
                            time.sleep(0.5)

                            # --- 3. Click Execution ---
                            submit_btn = wait.until(EC.presence_of_element_located((
                                By.XPATH, "//button[contains(normalize-space(.), 'Submit TB Report')]"
                            )))
                            driver.execute_script("arguments[0].click();", submit_btn)
                            print(f"Form submission executed successfully. (line: {sys._getframe().f_lineno})")


                            # 1. Explicitly wait until the SweetAlert confirm button is interactive on the screen viewport
                            yes_save_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
                            
                            # 2. Execute the click using JavaScript to guarantee execution through the backdrop fade overlay
                            driver.execute_script("arguments[0].click();", yes_save_btn)
                            print(f"Confirmation modal 'Yes, Save' button successfully clicked. (line: {sys._getframe().f_lineno})")

                            


                            # 1. Wait for and capture the SweetAlert2 duplicate container
                            # Now alert_container is actively utilized to locate elements relative to itself!
                            alert_container = WebDriverWait(driver, 1).until(
                                EC.visibility_of_element_located((By.CLASS_NAME, "swal2-container"))
                            )
                            mySleepFunction(1)

                            # 2. Scope the search INSIDE the alert_container for maximum speed and accuracy
                            error_text_element = alert_container.find_element(By.ID, "swal2-html-container")
                            mySleepFunction(2)

                            if "इस वर्ष और तिमाही" in error_text_element.text: # "recorded for this year/quarter"
                                print(f"Duplicate warning caught: {error_text_element.text}")
                                
                                # Target the confirm button inside the container element structure directly
                                ok_button = alert_container.find_element(By.CSS_SELECTOR, "button.swal2-confirm")
                                
                                js_force_click_and_close = """
                                    var btn = arguments[0];
                                    if (btn) {
                                        btn.focus();
                                        ['mousedown', 'mouseup', 'click'].forEach(function(evtName) {
                                            var e = new MouseEvent(evtName, { bubbles: true, cancelable: true, view: window });
                                            btn.dispatchEvent(e);
                                        });
                                        console.log("Synthesized interaction sequence dispatched to OK button.");
                                        
                                        setTimeout(function() {
                                            var container = document.querySelector(".swal2-container");
                                            if (container) {
                                                container.remove();
                                                document.body.classList.remove("swal2-shown", "swal2-height-auto");
                                                console.log("Forced dynamic cleanup of blocking SweetAlert layers completed.");
                                            }
                                        }, 300);
                                    }
                                """
                                driver.execute_script(js_force_click_and_close, ok_button)
                                
                                WebDriverWait(driver, 2).until(
                                    EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-container"))
                                )
                                
                                # 4. Trigger form clearing operations safely via JavaScript click
                                logging.info(f"Resetting form for next member: {card_number} (line: {sys._getframe().f_lineno})")
                                print(f"Resetting form for next member: {card_number} (line: {sys._getframe().f_lineno})")

                                # Locate the element context structure explicitly
                                specific_clear = WebDriverWait(driver, 2).until(
                                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'search-row-two')]//button[contains(., 'Clear')]"))
                                )
                                
                                # Execute script-based click to completely avoid Interception / State exceptions
                                driver.execute_script("arguments[0].click();", specific_clear)
                                print("Clear button successfully triggered via JS execution.")
                            else:
                                ok_button = alert_container.find_element(By.CSS_SELECTOR, "button.swal2-confirm")
                                
                                js_fallback_close = """
                                    var btn = arguments[0];
                                    if (btn) {
                                        btn.focus();
                                        ['mousedown', 'mouseup', 'click'].forEach(function(n) {
                                            btn.dispatchEvent(new MouseEvent(n, { bubbles: true, cancelable: true }));
                                        });
                                        setTimeout(function() {
                                            var c = document.querySelector(".swal2-container");
                                            if(c) { c.remove(); document.body.classList.remove("swal2-shown"); }
                                        }, 300);
                                    }
                                """
                                driver.execute_script(js_fallback_close, ok_button)
                                
                                WebDriverWait(driver, 5).until(
                                    EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-container"))
                                )
                                check_date_and_village_set()
























                            if i < total_buttons:
                                counter = counter + 1
                                continue

                        except Exception as e: # inner exception of option No, No, & photo upload
                            print(f"An error occurred during automation: {e} line: {sys._getframe().f_lineno}")

#####################################

                    except Exception as e:
                        print(f"Error clicking button #{i + 1}:line: {sys._getframe().f_lineno}")
                        
                        # Note: If clicking a button causes a full page reload or changes the DOM structure, 
                        # you will need to re-fetch the element list inside the loop to avoid StaleElementReferenceException.
                        
                    except Exception as e:
                        print(f"Could not click button #{1}:  line: {sys._getframe().f_lineno}")

                time.sleep(3)
                #Screening Logic will go here$$$$$$$$$$$$$$$$$$$$$$$$$$$$                

            except Exception as e: # exception of ration card for loop
                print(f"Pipeline crashed for card: {card_number} with error:  line: {sys._getframe().f_lineno}")
                

    except Exception as e: # exception of login function
        print(f"An error occurred while filling the form: line: {sys._getframe().f_lineno}")




# start executing
login()