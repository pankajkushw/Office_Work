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
import random as rand_mod



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
target_village = "Balrampur"
target_date = "10-08-2026" 
driver.implicitly_wait(10) 
def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)# Define an explicit wait timeout
wait = WebDriverWait(driver, 10)

# Configure the logging system
logging.basicConfig(
    filename='Balrampur_29102026.log',         # Name of the log file
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

def for_single_entry():
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


    try:
        # 1. Locate the Age input field using its formcontrolname attribute
        age_field = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='age_number']"))
        )
        age_value_str = age_field.get_attribute("value")
        
        if age_value_str and age_value_str.isdigit():
            age = int(age_value_str)
            
            # 2. Check if the age meets your condition (60 or above)
            if age >= 60:
                # UPDATED: Using formcontrolname="name" as shown in the DOM snippet
                name_field = driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='name']")
                member_name = name_field.get_attribute("value")
                print(f"[SKIP] Member '{member_name}' with Ration Card '{card_number}' is {age} years old. Moving to next member.")
                # Action: Place your loop's 'continue' or next-member logic here
                # Locate the element context structure explicitly
                specific_clear = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'search-row-two')]//button[contains(., 'Clear')]"))
                )
                
                # Execute script-based click to completely avoid Interception / State exceptions
                driver.execute_script("arguments[0].click();", specific_clear)
                print("Clear button successfully triggered via JS execution.")

    except Exception as e:
        print(f"Failed to verify member age criteria: {e}")

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
    
    ok_button = WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, ".swal2-confirm"))
    )
    
    ok_button.click()
    # Locate the element context structure explicitly
    specific_clear = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'search-row-two')]//button[contains(., 'Clear')]"))
    )

    # Execute script-based click to completely avoid Interception / State exceptions
    driver.execute_script("arguments[0].click();", specific_clear)
    mySleepFunction(3)
    print("Clear button successfully triggered from First_entry().")
    check_date_and_village_set()



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

        # 1. Fill Date Input via JavaScript
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
        # Click the dropdown trigger container to open the option panel
        village_dropdown = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.ID, "mat-select-8"))
        )
        village_dropdown.click()

        # Wait for the option element to be present and visible in the DOM
        # (Fixed typo from 'visibility_of_element_factor_located' to 'visibility_of_element_located')
        village_option = WebDriverWait(driver, 2).until(
            EC.visibility_of_element_located((By.XPATH, f"//mat-option//span[@class='mat-option-text' and contains(text(), '{target_village}')]"))
        )

        # Use a JavaScript Click to bypass any animation overlay blocks
        driver.execute_script("arguments[0].click();", village_option)

        # 3. Click the Continue Button
        continue_button = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continue')]"))
        )
        continue_button.click()


        # 1. Click the Select Year dropdown trigger container using formcontrolname
        year_dropdown = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "mat-select[formcontrolname='screening_year']"))
        )
        year_dropdown.click()

        # 2. Wait for the overlay option matching "2026" to be visible
        year_option = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//mat-option//span[@class='mat-option-text' and normalize-space(text())='2026']"))
        )

        # 3. Perform a JavaScript click to bypass any animation overlay delays
        driver.execute_script("arguments[0].click();", year_option)

        # 1. Click the Quarter dropdown trigger container using formcontrolname
        quarter_dropdown = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "mat-select[formcontrolname='screening_quarter']"))
        )
        quarter_dropdown.click()

        # 2. Wait for the overlay option matching "Quarter 3 (Jul-Sep)" to be visible
        quarter_option = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//mat-option//span[@class='mat-option-text' and normalize-space(text())='Quarter 3 (Jul-Sep)']"))
        )

        # 3. Perform a JavaScript click to bypass any animation overlay delays
        driver.execute_script("arguments[0].click();", quarter_option)
        mySleepFunction(3)

    else:
        print(f"Not able to set date and Village in check_and_set_date_and_village() {sys._getframe().f_lineno})")
        return False


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
        ration_cards = ['226483576521', '226483601097', '226483667049', '226483703830', '226483755810', '226483769625', '226483774550', '226484119669', '226484178179', '226484206594', '226484562465', '226484596878', '226484695880', '226484739458', '226484825381', '226484825381', '226484833370', '226484979270', '226485094395', '226485131056', '226485173152', '226485251662', '226485304474', '226485300070', '226485308033', '226485343682', '226485356325', '226485407248', '226485415943', '226485440211', '226485472365', '226485518572', '226485532915', '226485565646', '226485700052', '226485723391', '226485761665', '226485785097', '226485789611', '226485799620', '226485822795', '226485895511', '226485928338', '226485937377', '226485973486', '226485986138', '226486010735', '226486023898', '226486030803', '226486045276', '226486053189', '226486167483', '226486263620', '226486346401', '226486364081', '226486398109', '226486462413', '226486464833', '226486468809', '226486501767', '226486583090', '226486592494', '226486608138', '226486670707', '226486735944', '226486898568', '226486936527', '226486972192', '226487029294', '226487079079', '226487111348', '226487180090', '226487287364', '226487299600', '226487369599', '226487380174', '226487502188', '226487542251', '226487578025', '226487670895', '226487698225', '226487745368', '226487862316', '226487904646', '226487927872', '226488008757', '226488015398', '226488044714', '226488084566', '226488150138', '226488190761', '226488198509', '226488233248', '226488271634', '226488314150', '226488405262', '226488405692', '226488556103', '226488718561', '226488734908', '226488753497', '226488770141', '226488774372', '226488790378', '226488844791', '226488860367', '226488908355', '226488947886', '226489094621', '226489176158', '226489230437', '226489296195', '226489422960', '226489475586', '226489498851', '226489501599', '226489548065', '226489588311', '226489609237', '226489620354', '226489624483', '226489635285', '226489640402', '226489702177', '226489707568', '226489743106', '226489752743', '226489801440', '226489826398', '226489979539', '226489995456', '226481156300', '226481280950', '226481336256', '226481749253', '226481996251', '226482046050', '226482145497', '226482196630', '226482262496', '226482325966', '226482519894', '226482592397', '226482988083', '226483374901', '226483460877', '226483561427', '226483996236', '226484056879', '226484114072', '226484391118', '226484585766', '226484614790', '226484829210', '226484837424', '226485036739', '226485360218', '226485604845', '226485717252', '226485744797', '226485759963', '226485780226', '226485832537', '226486034146', '226486053497', '226486058396', '226486117845', '226486128071', '226486512983', '226486529259', '226486539494', '226486636987', '226486724756', '226486819896', '226486962512', '226487078252', '226487172375', '226487310928', '226487413804', '226487632111', '226487691950', '226487709487', '226487848264', '226488143666', '226488227194', '226488570440', '226488671751', '226488887164', '226488902968', '226488947325', '226489254175', '226489266146', '226489336251', '226489587274', '226489665424', '226489700178', '226489709765', '226489813931', '226489913534', '226484565907', '226480191911', '226480198862', '226480291170', '226480521721', '226480904942', '226481132806', '226481132835', '226481141190', '226481414157', '226481445711', '226481453746', '226481500292', '226481544141', '226481733813', '226482060613', '226482745057', '226482745796', '226482785871', '226482993304', '226483023667', '226483059864', '226483064004', '226483154779']
        
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


                popups = driver.find_elements(By.CLASS_NAME, "swal2-container")
                time.sleep(0.5)  # Allow the DOM to stabilize after the searc
                if len(popups) > 0 and popups[0].is_displayed():
                    # POPUP APPEARED: Locate the confirm button and click it using JavaScript execution
                    ok_button = driver.find_element(By.CSS_SELECTOR, "button.swal2-confirm")
                    driver.execute_script("arguments[0].click();", ok_button)
                    print("Condition Met: Popup appeared and was successfully closed.")
                    continue
                else:
                    # NO POPUP: Proceed normally with your script sequence
                    print("Condition Met: No popup found. Proceeding with regular operations.")




                identification_header = driver.find_elements(By.XPATH, "//*[contains(text(), 'Identification Details')]")
                time.sleep(0.5)  # Allow the DOM to stabilize after the search
                if len(identification_header) > 0:
                    print(f"Identification Details section is already visible on the screen: {card_number} line: {sys._getframe().f_lineno}")
                    logging.info("Bypassed select buttons loop because Identification Details container is active.")
                    for_single_entry()
                    continue  # Skip to the next card number in the outer loop


                print(f"Search Member of Ration Card {card_number}. (line: {sys._getframe().f_lineno})")
                #Screening Logic will go here$$$$$$$$$$$$$$$$$$$$$$$$$$$
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "custom-table")))
                # Target the buttons precisely using the class name 'action-btn' shown in your HTML
                select_buttons = driver.find_elements(By.CSS_SELECTOR, "button.action-btn")
                total_buttons = len(select_buttons)

                print(f"Found {total_buttons} matching 'Select' buttons in {card_number}. line: {sys._getframe().f_lineno}")
                logging.info(f"Member in :  {card_number} : {total_buttons}")

                counter = 0

                
                for i in range(total_buttons):
                    print(f"Processing Select button #{i + 1} of {total_buttons} for card {card_number}. line: {sys._getframe().f_lineno}")
                    
                    try:
                        #Second time date of visit, village has to be set and search again
                        if counter > 0:
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
                        print(f"Proceeding data update... (line: {sys._getframe().f_lineno})")

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


                            try:
                                # 1. Locate the Age input field using its formcontrolname attribute
                                age_field = WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='age_number']"))
                                )
                                age_value_str = age_field.get_attribute("value")
                                
                                if age_value_str and age_value_str.isdigit():
                                    age = int(age_value_str)
                                    
                                    # 2. Check if the age meets your condition (60 or above)
                                    if age >= 60:
                                        # UPDATED: Using formcontrolname="name" as shown in the DOM snippet
                                        name_field = driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='name']")
                                        member_name = name_field.get_attribute("value")
                                        print(f"[SKIP] Member '{member_name}' with Ration Card '{card_number}' is {age} years old. Moving to next member.")
                                        # Action: Place your loop's 'continue' or next-member logic here
                                        # Locate the element context structure explicitly
                                        specific_clear = WebDriverWait(driver, 2).until(
                                            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'search-row-two')]//button[contains(., 'Clear')]"))
                                        )
                                        mySleepFunction(3)   
                                        # Execute script-based click to completely avoid Interception / State exceptions
                                        driver.execute_script("arguments[0].click();", specific_clear)
                                        print("Clear button successfully triggered via JS execution.")

                                        counter += 1
                                        continue  # Skip to the next iteration of the outer loop

                            except Exception as e:
                                print(f"Failed to verify member age criteria: {e}")



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
                            time.sleep(0.5)
                            yes_save_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
                            # 2. Execute the click using JavaScript to guarantee execution through the backdrop fade overlay
                            driver.execute_script("arguments[0].click();", yes_save_btn)
                            print(f"Confirmation modal 'Yes, Save' button successfully clicked. (line: {sys._getframe().f_lineno})")
                            mySleepFunction(2)
                            


                            # 1. Wait for and capture the SweetAlert2 duplicate container
                            # Now alert_container is actively utilized to locate elements relative to itself!
                            alert_container = WebDriverWait(driver, 3).until(
                                EC.visibility_of_element_located((By.CLASS_NAME, "swal2-container"))
                            )

                            # 2. Scope the search INSIDE the alert_container for maximum speed and accuracy
                            error_text_element = alert_container.find_element(By.ID, "swal2-html-container")

                            if "इस वर्ष और तिमाही" in error_text_element.text: # "recorded for this year/quarter"
                                print(f"Duplicate warning caught: {error_text_element.text}")
                                
                                # Target the confirm button inside the container element structure directly
                                ok_button = alert_container.find_element(By.CSS_SELECTOR, "button.swal2-confirm")

                                js_force_click_and_close = """
                                    var btn = arguments[0]; // Fixed: Explicitly target index 0
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
                                        }, 150); // Reduced delay for faster UI convergence
                                    }
                                """
                                # Execute the corrected script injection pass
                                
                                driver.execute_script(js_force_click_and_close, ok_button)


                                logging.info(f"Resetting form for next member: {card_number} (line: {sys._getframe().f_lineno})")
                                print(f"Resetting form for next member: {card_number} (line: {sys._getframe().f_lineno})")

                                try:
                                    # 1. Target the button tag specifically inside the structural class layout container
                                    specific_clear = WebDriverWait(driver, 5).until(
                                        EC.presence_of_element_located((
                                            By.XPATH, 
                                            "//div[contains(@class, 'search-row-two')]//button[contains(., 'Clear')]"
                                        ))
                                    )
                                    
                                    time.sleep(3) 
                                    
                                    # 2. FIXED: Array index context added to arguments object to successfully fire the action
                                    driver.execute_script("arguments[0].click();", specific_clear)
                                    print("Clear button triggered successfully via JavaScript execution.")
                                    
                                except Exception as e:
                                    print(f"Failed to trigger the Clear button: {e}")
                                    specific_clear_span = WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((
                                        By.XPATH, 
                                        "//div[contains(@class, 'search-row-two')]//span[@class='mat-button-wrapper' and contains(text(), 'Clear')]"
                                    ))
                                    )
                                    driver.execute_script("arguments[0].click();", specific_clear_span)

                                
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