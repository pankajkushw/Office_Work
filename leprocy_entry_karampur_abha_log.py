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
target_village = "Karampur"
target_date = "01-09-2026" 
driver.implicitly_wait(10) 
def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)# Define an explicit wait timeout
wait = WebDriverWait(driver, 10)

# Configure the logging system
logging.basicConfig(
    filename='Karampur_Abha_ID.log',         # Name of the log file
    filemode='a',              # 'a' to append logs, 'w' to overwrite each run
    format='%(asctime)s - %(levelname)s - %(message)s', # Log structure
    level=logging.INFO         # Capture INFO level messages and above
)

def abha_member_log(card_number, message):
    # Locate the name input field safely using its formcontrolname attribute
    name_input = wait.until(
        EC.presence_of_element_located((By.XPATH, "//input[@formcontrolname='name']"))
    )
    
    # Retrieve the text present inside the input field value attribute
    name_value = name_input.get_attribute("value")
    logging.info(f"Card Number: {card_number} | {message} for Name: {name_value}")
    print(f"Retrieved Name: {name_value}")






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

    # Check if both fields contain the Angular invalid class marker
    if "ng-invalid" in date_classes and "ng-invalid" in village_classes:
        print(f"Both fields are highlighted in red (invalid status). Injecting target data... (Line: {sys._getframe().f_lineno})")
        # 1. Wait for global Angular loader to disappear completely
        wait.until(EC.invisibility_of_element_located((By.TAG_NAME, "app-loader")))
        
        # 2. Input Planned Visit Date via JavaScript execution
        # (This bypasses click interceptions if the read-only or overlay blocks standard input typing)
        date_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Visit Date')] | //mat-form-field[contains(., 'Visit Date')]//input")))
        driver.execute_script("arguments[0].value = arguments[1];", date_input, target_date)
        # Trigger standard input/change events so Angular detects the value update
        driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", date_input)
        #driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: True }));", date_input)
        print(f"Set planned visit date to: {target_date} (Line: {sys._getframe().f_lineno})")

        # # # 3. Open Village Dropdown Panel
        # print(f"Attempting to select village: {target_village} (Line: {sys._getframe().f_lineno})")
        # village_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-select | //mat-form-field[contains(., 'Village')]//mat-select")))
        # village_dropdown.click()
        # print(f"Village dropdown opened (Line: {sys._getframe().f_lineno})")

        # # 4. Handle Option Selection using your expanded XPATH pattern with case-insensitive matching
        # print(f"Attempting to select village: {target_village} (Line: {sys._getframe().f_lineno})")
        # option_xpath = f"//mat-option[contains(translate(., 'BALRAMPUR', 'Balrampur'), '{target_village.lower()}')] | //mat-option//span[contains(translate(text(), 'BALRAMPUR', 'Balrampur'), '{target_village.lower()}')]"
        # option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
        # option.click()
        # print(f"Selected village: {target_village} line: {sys._getframe().f_lineno}")

        # 1. Locate and click the Angular dropdown menu container
        dropdown_xpath = "//mat-select[contains(., 'Select Village')] | //div[contains(text(), 'Select Village')] | //mat-form-field[contains(., 'Select Village')]"
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
        dropdown.click()
        print(f"Dropdown opened.line: {sys._getframe().f_lineno}")

        # 2. Wait for the overlay option panel to pop up and click the targeted choice
        option_xpath = f"//mat-option[contains(., '{target_village}')] | //span[contains(@class, 'mat-option-text') and contains(text(), '{target_village}')]"
        option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
        option.click()
        print(f"Selected village: {target_village} (line: {sys._getframe().f_lineno})")
        
        # 3. Locate and click the "Continue" button
        # Using text matching helps isolate it even if it changes from disabled to active state
        continue_btn_xpath = "//button[contains(., 'Continue')] | //span[contains(text(), 'Continue')]/ancestor::button"
        continue_btn = wait.until(EC.element_to_be_clickable((By.XPATH, continue_btn_xpath)))
        continue_btn.click()
        print(f"Successfully clicked 'Continue'. (line: {sys._getframe().f_lineno})")

        # # 1. Broadly target the search button card using text and style classes
        # search_btn_xpath = (
        #     "//button[@type='submit' or contains(., 'Continue')]"
        #     " | //mat-form-field//following::button[contains(., 'Continue')]"
        #     " | //span[contains(text(), 'Continue')]/ancestor::button"
        #     " | //button[contains(@class, 'mat-focus-indicator') and contains(., 'Continue')]"
        # )
        
        # # 2. Wait until the button is present in the DOM layout
        # search_btn = wait.until(EC.presence_of_element_located((By.XPATH, search_btn_xpath)))
        
        # # Try a standard driver click first to allow Angular event bubbles to fire naturally
        # search_btn.click()
        # print("Invalidated, setting again.")

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
            By.XPATH, "//*[normalize-space(.)='Go to Leprosy Abhiyan']"
        )))

        # Force the click via JavaScript to completely bypass any Angular Material 
        # click interception layers or animation delays.
        driver.execute_script("arguments[0].click();", leprosy_program_btn)


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
        print("Successfully selected September 2nd.")

        # Locates the "Go to Entry Page" button using text-based matching
        entry_page_xpath = (
            "//button[contains(., 'Go to Entry Page')] | "
            "//a[contains(., 'Go to Entry Page')] | "
            "//*[text()='Go to Entry Page']"
        )
        
        # Wait up to 10 seconds for the button to be visible and clickable
        entry_page_btn = wait.until(EC.element_to_be_clickable((By.XPATH, entry_page_xpath)))
        entry_page_btn.click()
        print(f"Successfully clicked 'Go to Entry Page' button.line: {sys._getframe().f_lineno}")




        # 1. Locate and click the Angular dropdown menu container
        dropdown_xpath = "//mat-select[contains(., 'Select Village')] | //div[contains(text(), 'Select Village')] | //mat-form-field[contains(., 'Select Village')]"
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
        dropdown.click()
        print(f"Dropdown opened.line: {sys._getframe().f_lineno}")

        # 2. Wait for the overlay option panel to pop up and click the targeted choice
        option_xpath = f"//mat-option[contains(., '{target_village}')] | //span[contains(@class, 'mat-option-text') and contains(text(), '{target_village}')]"
        option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
        option.click()
        print(f"Selected village: {target_village} (line: {sys._getframe().f_lineno})")
        
        # 3. Locate and click the "Continue" button
        # Using text matching helps isolate it even if it changes from disabled to active state
        continue_btn_xpath = "//button[contains(., 'Continue')] | //span[contains(text(), 'Continue')]/ancestor::button"
        continue_btn = wait.until(EC.element_to_be_clickable((By.XPATH, continue_btn_xpath)))
        continue_btn.click()
        print(f"Successfully clicked 'Continue'. (line: {sys._getframe().f_lineno})")

        # Replace this list with your actual target Ration Card numbers
        ration_cards = ['226481549416', '226481558141', '226481582003', '226481676485', '226481720534', '226481876031', '226481925498', '226481982565', '226482058325', '226482138402', '226482162771', '226482290791', '226482339614', '226482348635', '226482367487', '226482385216', '226482395342', '226482410192', '226482485951', '226482677143', '226482724449', '226482775958', '226482830024', '226482852798', '226482859226', '226482860037', '226482997102', '226483145833', '226483152964', '226483403516', '226483467766', '226483621017', '226483700704', '226483796604', '226483911461', '226484121296', '226484296040', '226484372945', '226484451179', '226484454205', '226484484694', '226484545256', '226484555413', '226484556390', '226484633097', '226484683697', '226484821848', '226484826512', '226484847690', '226484964887', '226484980029', '226485024136', '226485256604', '226485275322', '226485322819', '226485494854', '226485530930', '226485689211', '226485833181', '226485838143', '226485914042', '226485975569', '226486074538', '226486085458', '226486172417', '226486410480', '226486541945', '226486545611', '226486611961', '226486632384', '226486836326', '226486840793', '226487022675', '226487044618', '226487119281', '226487211245', '226487231663', '226487473528', '226487488043', '226487523178', '226487560949', '226487563632', '226487575193', '226487592048', '226487729536', '226487745407', '226487813609', '226487938307', '226488029265', '226488133132', '226488174897', '226488195579', '226488219524', '226488230810', '226488526423', '226488531685', '226488653211', '226488682017', '226488722621', '226489063589', '226489116097', '226489160557', '226489209726', '226489266775', '226489290563', '226489404669', '226489486092', '226489522904', '226489619845', '226489645995', '226489804594', '226489938948', '226489946464', '223890581182', '223894092395', '223895965824', '226480005419', '226480012452', '226480025106', '226480055484', '226480074680', '226480101903', '226480105323', '226480174214', '226480193200', '226480209774', '226480287165', '226480330956', '226480332827', '226480361280', '226480400616', '226480445084', '226480447201', '226480604572', '226480607683', '226480637790', '226480664676', '226480696369', '226480767260', '226480767260', '226480770516', '226480880658', '226480934863', '226480951560', '226480961433', '226480969484', '226481004451', '226481045784', '226481049249', '226481056826', '226481090303', '226481156601', '226481169061', '226481188209', '226481200170', '226481238219', '226481239401', '226481245044', '226481290385', '226481303408', '226481356867', '226481360371', '226481457944', '226481467957', '226481490864', '226481493179', '226481514922', '226481536044', '226481559548', '226481559905', '226481572522', '226481613763', '226481627295', '226481640989', '226481640989', '226481674926', '226481683115', '226481706919', '226481709808', '226481712519', '226481737352', '226481749700', '226481764705', '226481773972', '226481816143', '226481820767', '226481842580', '226482023164', '226482030415', '226482116836', '226482123273', '226482126340', '226482205837', '226482220868', '226482229360', '226482249974', '226482271868', '226482312668', '226482354202', '226482410133', '226482410148', '226482429095', '226482442716', '226482446124', '226482489495', '226482492237', '226482532452', '226482532452', '226482592928', '226482595715', '226482599014', '226482602267', '226482617419', '226482645899', '226482731139', '226482768277', '226482779193', '226482781891', '226482838050', '226482842397', '226482884810', '226482908044', '226482921656', '226482929254', '226482978952', '226482994402', '226483003596', '226483022279', '226483033214', '226483058828', '226483142975', '226483183391', '226483201549', '226483232652', '226483261145', '226483279462', '226483298701', '226483298701', '226483366042', '226483417066', '226483446389', '226483453664', '226483467047', '226483474783', '226483477106', '226483604150', '226483639993', '226483664146', '226483667422', '226483700202', '226483713911', '226483718016', '226483755056', '226483775106', '226483776556', '226483806693', '226483915852', '226483921048', '226483952582', '226484010409', '226484057600', '226484057690', '226484061290', '226484063450', '226484071950', '226484086981', '226484088411', '226484148252', '226484175652', '226484179416', '226484180012', '226484182575', '226484188453', '226484209477', '226484218419', '226484227428', '226484229421', '226484249632', '226484277454', '226484299897', '226484313212', '226484333344', '226484369874', '226484370730', '226484379832', '226484385888', '226484429089', '226484431178', '226484431178', '226484439035', '226484496229', '226484514048', '226484559978', '226484592764', '226484596459', '226484598872', '226484600331', '226484603056', '226484614001', '226484619650', '226484670816', '226484687845', '226484801193', '226484805330', '226484837822', '226484839152', '226484874207', '226484876228', '226484893666', '226484947271', '226484955045', '226485010295', '226485013800', '226485018301', '226485028154', '226485035710', '226485039637', '226485207575', '226485219467', '226485247032', '226485261817', '226485279516', '226485330192', '226485430050', '226485463425', '226485497007', '226485501706', '226485513335', '226485519298', '226485546385', '226485580657', '226485584546', '226485602334', '226485633582', '226485645052', '226485661404', '226485753550', '226485791784', '226485826553', '226485874066', '226485880098', '226485901869', '226485930675', '226486034718', '226486081780', '226486085989', '226486108346', '226486114296', '226486116593', '226486142492', '226486160961', '226486177064', '226486194798', '226486197249', '226486203315', '226486244366', '226486270421', '226486275283', '226486332843', '226486350073', '226486358704', '226486383751', '226486390263', '226486390263', '226486403118', '226486403744', '226486404133', '226486424488', '226486477743', '226486487751', '226486503826', '226486521364', '226486530532', '226486542122', '226486556894', '226486584085', '226486584085', '226486673137', '226486682564', '226486696586', '226486714524', '226486745487', '226486769022', '226486788949', '226486834342', '226486840355', '226486959555', '226486974434', '226486984823', '226487023747', '226487039646', '226487074070', '226487076227', '226487095134', '226487109465', '226487121064', '226487161387', '226487176976', '226487190300', '226487193914', '226487272760', '226487276031', '226487301242', '226487325773', '226487353299', '226487376853', '226487417410', '226487453230', '226487540391', '226487564697', '226487571909', '226487591812', '226487632740', '226487655213', '226487705397', '226487707809', '226487775645', '226487804978', '226487846938', '226487849592', '226487878970', '226487919888', '226487923313', '226487932918', '226487978667', '226488066768', '226488100552', '226488192724', '226488206098', '226488212849', '226488267470', '226488278049', '226488363287', '226488466477', '226488522164', '226488524446', '226488530628', '226488531971', '226488536337', '226488553581', '226488571491', '226488599210', '226488605255', '226488627867', '226488638425', '226488667057', '226488683672', '226488686448', '226488704366', '226488757370', '226488796887', '226488907313', '226488983851', '226489036405', '226489055671', '226489069015', '226489090524', '226489176840', '226489180457', '226489197725', '226489218964', '226489253499', '226489254753', '226489417513', '226489529321', '226489539898', '226489561982', '226489564289', '226489584256', '226489586112', '226489646226', '226489679133', '226489702521', '226489736524', '226489776416', '226489795683', '226489826479', '226489878451', '226489903241', '226489917157', '226489984089', '226481031582', '226481086311', '226481246697', '226481287369', '226481295091', '226481329177', '226481370727', '226481445235', '226481475437', '226481490056', '226481498137', '226481538910', '226481568664', '226481731382', '226481767777', '226481783742', '226481816740', '226481914283', '226481931836', '226482100875', '226482205405', '226482214647', '226482219572', '226482331544', '226482465109', '226482473179', '226482486773', '226482597161', '226482597425', '226482648129', '226482662478', '226482714442', '226482748625', '226482834104', '226482918139', '226483056531', '226483111689', '226483187660', '226483197585', '226483346894', '226483360580', '226483400789', '226483585926', '226483646792', '226483652305', '226483661577', '226483783307', '226483783307', '226484272146', '226484319895', '226484355015', '226484394012', '226484530730', '226484597586', '226484643774', '226484663169', '226484776610', '226484842055', '226485316537', '226485593171', '226485593536', '226485632613', '226485706648', '226485732124', '226485747657', '226485821857', '226485851903', '226485879336', '226485991440', '226486053393', '226486077050', '226486162149', '226486247960', '226486260561', '226486261393', '226486397072', '226486515964', '226486535743', '226486580336', '226486632946', '226486668463', '226486799132', '226486849531', '226486853603', '226486873055', '226486938586', '226486943358', '226487015741', '226487080693', '226487102828', '226487197251', '226487202125', '226487253122', '226487256814', '226487388396', '226487447259', '226487591416', '226487626974', '226487650645', '226487666863', '226487678954', '226487742085', '226487762811', '226487805756', '226487864158', '226487893332', '226487990250', '226488039906', '226488053587', '226488085912', '226488093473', '226488106469', '226488122234', '226488123151', '226488266071', '226488311893', '226488425215', '226488549478', '226488613216', '226488617835', '226488658989','226488752364', '226488772008', '226489043496', '226489082771', '226489316989', '226489395447', '226489512586', '226489616105', '226489710028', '226489742367', '226489807035', '226489843278', '226489871259', '226489919266', '226489928354', '226489993637', '226489452114']
        round_complete = False
        for card_number in ration_cards:
            
            ret = check_date_and_village_set()
            print(f"Executing sequence for card entry: {card_number}")
            logging.info(f"Opening:  {card_number}")
            try:
                # Targets the input field directly associated with the ID card icon,
                # explicitly avoiding any date picker fields containing calendar icons.
                if (ret == False):
                    print(f"Date and Village fields were not set correctly. Re-injecting values for card: {card_number} (line: {sys._getframe().f_lineno})")
                    #check_date_and_village_set()
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

                # Try a standard driver click first to allow Angular event bubbles to fire naturally
                search_btn.click()

                # # Wait for the "OK" button to be clickable and then click it
                # ok_button = WebDriverWait(driver, 10).until(
                #     EC.element_to_be_clickable((By.CLASS_NAME, "swal2-confirm"))
                # )
                # ok_button.click()

                print(f"Search Member of Ration Card. (line: {sys._getframe().f_lineno})")

                #Screening Logic will go here$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                # Find all active "Select" buttons in the table
                # This uses a partial text match or exact match on the text inside the button/link
             
                try:
                    # 1. Check if the "Ayushman Card Not Found" popup appears (waits up to 3 seconds)
                    popup_header = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Ayushman Card Not Found')]"))
                    )
                    
                    print(f"Popup encountered: Ayushman Card Not Found for card {card_number}. Dismissing and skipping. (line: {sys._getframe().f_lineno})")
                    abha_member_log(card_number, "Ayushman Card Not Found")
                    #logging.warning(f"Ayushman Card Not Found for card: {card_number}")

                    # 2. Locate and click the 'OK' button to dismiss the modal
                    # Using a text match on the purple button text 'OK'
                    ok_button = driver.find_element(By.XPATH, "//button[contains(text(), 'OK')] | //*[text()='OK']")
                    ok_button.click()
                    
                    # 3. Wait briefly for the modal backdrop overlay to disappear before continuing the loop
                    WebDriverWait(driver, 3).until(EC.staleness_of(popup_header))
                    
                    # 4. Skip the rest of the current iteration and move to the next member/card
                    continue

                except TimeoutException:
                    # No popup appeared within 3 seconds, proceed normally with the form actions
                    pass

                identification_header = driver.find_elements(By.XPATH, "//*[contains(text(), 'Identification Details')]")
                if len(identification_header) > 0:
                    print(f"Identification Details section is already visible on the screen: {card_number} line: {sys._getframe().f_lineno}")
                    logging.info("Bypassed select buttons loop because Identification Details container is active.")
                    
                    for_single()
                    
                    try:
                        abha_field = driver.find_element(By.XPATH, "//mat-form-field[contains(., 'Abha Id')]//input")
                        abha_id_value = abha_field.get_attribute("value")
                        print(f"Abha ID found for card {card_number}: {abha_id_value} line: {sys._getframe().f_lineno}")
                        # # CHANGE HERE: Skip the loop if it is NOT present
                        # if not abha_id_value or not abha_id_value.strip():
                        #     print(f"Abha ID is missing for card {card_number}. Skipping next steps.")
                        #     continue  
                            
                    except Exception as e:
                        #logging.warning(f"Could not read Abha ID field: {str(e)}")
                        # If the field can't be found, it's not present, so we skip
                        #print(f"Could not read Abha ID field for card {card_number}. Skipping next steps. Error: {str(e)} line: {sys._getframe().f_lineno}")
                        # Locate the name input field safely using its formcontrolname attribute
                        abha_member_log(card_number, 'Abha ID Not Found')
                else:
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
                            #1. Wait for global Angular loader to disappear completely
                            wait.until(EC.invisibility_of_element_located((By.TAG_NAME, "app-loader")))

                            # 2. Input Planned Visit Date via JavaScript execution
                            # (This bypasses click interceptions if the read-only or overlay blocks standard input typing)
                            date_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Visit Date')] | //mat-form-field[contains(., 'Visit Date')]//input")))
                            driver.execute_script("arguments[0].value = arguments[1];", date_input, target_date)
                            # Trigger standard input/change events so Angular detects the value update
                            driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", date_input)
                            #driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: True }));", date_input)
                            print(f"Set planned visit date to: {target_date}")

                            # # 4. Handle Option Selection using your expanded XPATH pattern with case-insensitive matching
                            # option_xpath = f"//mat-option[contains(translate(., 'BALRAMPUR', 'Balrampur'), '{target_village.lower()}')] | //mat-option//span[contains(translate(text(), 'BALRAMPUR', 'Balrampur'), '{target_village.lower()}')]"
                            # option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
                            # option.click()
                            # print(f"Selected village: {target_village} line: {sys._getframe().f_lineno}")
                            
                            # # # 3. Open Village Dropdown Panel
                            # print(f"Attempting to select village: {target_village} (Line: {sys._getframe().f_lineno})")
                            # village_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-select | //mat-form-field[contains(., 'Village')]//mat-select")))
                            # village_dropdown.click()
                            # print(f"Village dropdown opened (Line: {sys._getframe().f_lineno})")

                            # # 4. Handle Option Selection using your expanded XPATH pattern with case-insensitive matching
                            # print(f"Attempting to select village: {target_village} (Line: {sys._getframe().f_lineno})")
                            # option_xpath = f"//mat-option[contains(translate(., 'BALRAMPUR', 'Balrampur'), '{target_village.lower()}')] | //mat-option//span[contains(translate(text(), 'BALRAMPUR', 'Balrampur'), '{target_village.lower()}')]"
                            # option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
                            # option.click()
                            # print(f"Selected village: {target_village} line: {sys._getframe().f_lineno}")
                            UPPER = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                            LOWER = "abcdefghijklmnopqrstuvwxyz"
                            target_lower = target_village.lower()

                            try:
                                # 1. CLICK THE DROPDOWN CONTAINER TO OPEN IT
                                # Targets the exact mat-select element visible in your DOM panel
                                dropdown_xpath = "//mat-select[@formcontrolname='visite_village_code']"
                                dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
                                
                                # Scroll the dropdown container into view first and click it
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown)
                                time.sleep(0.5) # Brief pause for layout stability
                                dropdown.click()
                                
                                # 2. WAIT FOR THE OPTIONS OVERLAY TO APPAER AND SELECT THE VILLAGE
                                # Uses normalize-space(.) instead of text() to ignore layout tags
                                option_xpath = (
                                    f"//mat-option[contains(translate(normalize-space(.), '{UPPER}', '{LOWER}'), '{target_lower}')] | "
                                    f"//mat-option//span[contains(translate(normalize-space(.), '{UPPER}', '{LOWER}'), '{target_lower}')]"
                                )
                                
                                # Wait until the option is fully clickable in the newly opened overlay panel
                                option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
                                
                                # Scroll the target option into view (Crucial for long dropdown lists)
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", option)
                                time.sleep(0.3)
                                
                                # Click the option to properly satisfy Angular's form validation
                                option.click()
                                print(f"Successfully selected village: {target_village} line: {sys._getframe().f_lineno}")

                            except Exception as e:
                                print(f"Failed to select village '{target_village}' at line {sys._getframe().f_lineno}")
                                #print(f"Error Details: {type(e).__name__} - {e}")
                                raise e
                            

                             # 1. Broadly target the search button card using text and style classes
                            search_btn_xpath = (
                                "//button[@type='submit' or contains(., 'Continue')]"
                                " | //mat-form-field//following::button[contains(., 'Continue')]"
                                " | //span[contains(text(), 'Continue')]/ancestor::button"
                                " | //button[contains(@class, 'mat-focus-indicator') and contains(., 'Continue')]"
                            )
                            
                            # 2. Wait until the button is present in the DOM layout
                            search_btn = wait.until(EC.presence_of_element_located((By.XPATH, search_btn_xpath)))
                            
                            # Try a standard driver click first to allow Angular event bubbles to fire naturally
                            search_btn.click()

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

                            # Try a standard driver click first to allow Angular event bubbles to fire naturally
                            search_btn.click()
                            print(f"Standard browser search button click executed. (line: {sys._getframe().f_lineno})")




                        # Re-fetch the elements inside the loop to ensure they are fresh
                        buttons = driver.find_elements(By.CSS_SELECTOR, "button.action-btn")
                        current_button = buttons[i]
                        
                        # Scroll the specific button into view before interacting
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", current_button)
                        time.sleep(0.5)
                        
                        print(f"Clicking Select button #{i + 1}... (line: {sys._getframe().f_lineno})")
                        
                        # Use a reliable JavaScript click to bypass overlapping Angular components or overlay layers
                        driver.execute_script("arguments[0].click();", current_button)

                        # Initialize the flag tracker at the start of each loop iteration
                        screening_already_done = False

                        try:
                            # 1. Target the explicit SweetAlert container that is open on your screen
                            print(f"Checking for active SweetAlert warning layout... (line: {sys._getframe().f_lineno})")
                            swal_modal = WebDriverWait(driver, 3).until( # from 5 to 10
                                EC.presence_of_element_located((By.CLASS_NAME, "swal2-modal"))
                            )
                            
                            # Verify the specific error message is inside the modal header
                            if "Screening Already Done" in swal_modal.text:
                                print(f"⚠️ Match found: 'Screening Already Done' alert verified. (line: {sys._getframe().f_lineno})")
                                
                                # 2. Locate the precise SweetAlert confirm button using its dedicated library class
                                ok_btn = WebDriverWait(driver, 3).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm"))
                                )
                                
                                # 3. Force click via JavaScript to bypass any backdrop focus lock layers
                                driver.execute_script("arguments[0].click();", ok_btn)
                                print(f"SweetAlert 'OK' button successfully clicked. (line: {sys._getframe().f_lineno})")
                                
                                # # 4. Wait for the SweetAlert dark backdrop container to leave the DOM hierarchy entirely
                                # WebDriverWait(driver, 3).until( # from 5 to 10
                                #     EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-container"))
                                # )
                                WebDriverWait(driver, 5).until(
                                    EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-modal"))
                                )




                                # 4. Wait for the SweetAlert dark backdrop container to leave the DOM hierarchy entirely
                                # swal_container = driver.find_element(By.CLASS_NAME, "swal2-container")
                                # WebDriverWait(driver, 5).until(EC.staleness_of(swal_container))

                                print(f"Modal faded out. Workspace cleared. (line: {sys._getframe().f_lineno})")
                                
                                # Flip our loop bypass flag to True
                                screening_already_done = True

                        except Exception as e:
                            # If the modal doesn't exist, this block catch routes seamlessly into the normal flow
                            print(f"No active validation alert container intercepted line: {sys._getframe().f_lineno}. Continuing with normal form actions...")

                        # --- THE CRITICAL CONDITIONAL SKIP ---
                        if screening_already_done:
                            print(f"Skipping remaining form fields. Routing directly back to the next loop iteration... (line: {sys._getframe().f_lineno})")
                            #logging.info(f"screening already done for {card_number}: {i}")
                            continue  # Breaks the current execution string and pulls the next record smoothly

                        # --- REST OF FORM SUBMISSION ROUTINE CONTINUES BELOW ---
                        print(f"Proceeding with normal report creation actions... (line: {sys._getframe().f_lineno})")

                        try:
                            abha_input_element = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, "//input[@formcontrolname='abhaId']"))
                            )
                            abha_id_value = abha_input_element.get_attribute("value")
                            
                            # # CHANGE HERE: Skip the loop if it is NOT present
                            # if not abha_id_value or not abha_id_value.strip():
                            #     print(f"Abha ID is missing for card {card_number}. Skipping next steps.")
                            #     clear_search_btn = driver.find_element(By.XPATH, "//button[contains(., 'Search')]/following-sibling::button[contains(., 'Clear')]")
                            #     clear_search_btn.click()
                            if abha_id_value and abha_id_value.strip():
                                print(f"Abha ID is populated: {abha_id_value}")
                                #logging.info(f"Abha ID found for member: {abha_id_value}")
                                # Add your loop skip or 'continue' logic here if needed
                                
                            else:
                                print("Abha ID field is empty/blank.")

                                abha_member_log(card_number, 'abha_id_not_found')

                                clear_search_btn = driver.find_element(By.XPATH, "//button[contains(., 'Search')]/following-sibling::button[contains(., 'Clear')]")
                                clear_search_btn.click()
                                search_btn_xpath = (
                                    "//button[@type='submit' or contains(., 'Search')]"
                                    " | //mat-form-field//following::button[contains(., 'Search')]"
                                    " | //span[contains(text(), 'Search')]/ancestor::button"
                                    " | //button[contains(@class, 'mat-focus-indicator') and contains(., 'Search')]"
                                )
                                
                                # 2. Wait until the button is present in the DOM layout
                                search_btn = wait.until(EC.presence_of_element_located((By.XPATH, search_btn_xpath)))
                                # 3. Clean click execution strategy
                                # Try a standard driver click first to allow Angular event bubbles to fire naturally
                                search_btn.click()
                                #logging.info("Abha ID is not populated. Proceeding with form execution.")
                                # Add your form entry/generation logic here
                                continue
                                
                        except Exception as e:
                            #print(f"Error while checking Abha ID field: {e} line: {sys._getframe().f_lineno}")
                            print(f"Abha ID field is not present on the form. Skipping next steps for card {card_number}. line: {sys._getframe().f_lineno}")
                            abha_member_log(card_number, 'abha_id_not_found')
                            # If the field can't be found, it's not present, so we skip


                        try:
                           # Wait up to 10 seconds for the SweetAlert OK button to be clickable
                            ok_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm"))
                            )
                            ok_button.click()
                            print("Popup closed successfully.")
                            continue
                        except Exception as e:
                            print(f"Failed to close popup: Already Exist  line: {sys._getframe().f_lineno}")

                        time.sleep(1) 
                        #####################################
                        try:
                            print(f"screening for {card_number}: {i}")
                            # ----------------------------------------------------
                            # 1. Select "No" in the "Belongs to PVTG Category" Dropdown
                            # ----------------------------------------------------
                            # Locate and click the PVTG dropdown container to expand the options
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

                            except Exception as e   :
                                # अगर element 3 सेकंड में नहीं मिलता, तो script बिना क्रैश हुए इसे छोड़ देगी
                                print(f"गर्भवती / स्तनपान कराने वाली dropdown स्क्रीन पर दिखाई नहीं दिया। आगे बढ़ रहे हैं... (line: {sys._getframe().f_lineno})")

                            
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
                            print(f"Placeholder photo successfully routed to file stream. (line: {sys._getframe().f_lineno})")

                            # --- 2. Corrected Synchronization Wait ---
                            # Use the global 'wait' object instance to keep timeout metrics uniform.
                            # This explicitly waits for the image rendering preview frame to appear in the container layout.
                            wait.until(EC.presence_of_element_located((
                                By.XPATH, "//div[contains(@class, 'image')]//img | //img[not(@id) and @src] | //*[contains(@class, 'preview')]"
                            )))
                            print(f"Form validation refreshed: Photo preview detected. (line: {sys._getframe().f_lineno})")

                            # --- 3. Click Execution ---
                            submit_btn = wait.until(EC.presence_of_element_located((
                                By.XPATH, "//button[contains(normalize-space(.), 'Submit Leprosy Report')]"
                            )))
                            driver.execute_script("arguments[0].click();", submit_btn)
                            print(f"Form submission executed successfully. (line: {sys._getframe().f_lineno})")


                            # 1. Explicitly wait until the SweetAlert confirm button is interactive on the screen viewport
                            yes_save_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
                            
                            # 2. Execute the click using JavaScript to guarantee execution through the backdrop fade overlay
                            driver.execute_script("arguments[0].click();", yes_save_btn)
                            print(f"Confirmation modal 'Yes, Save' button successfully clicked. (line: {sys._getframe().f_lineno})")

                            mySleepFunction(3)
                            try:
                                # 1. Target the button via its unique SweetAlert confirmation class
                                success_ok_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
                                
                                # 2. Fix: Corrected syntax using arguments[0] to run the native browser click track
                                driver.execute_script("arguments[0].click();", success_ok_btn)
                                print(f"Success popup 'OK' button clicked via corrected JS call. (line: {sys._getframe().f_lineno})")

                            except Exception:
                                # Fallback Option: If the overlay layer blocks it, move the real pointer directly to the center and click
                                print(f"JavaScript click fallback initiated... Attempting Actions API click for 'OK' button. (line: {sys._getframe().f_lineno})")
                                success_ok_btn = driver.find_element(By.CSS_SELECTOR, "button.swal2-confirm")
                                ActionChains(driver).move_to_element(success_ok_btn).click().perform()
                                print(f"Success popup 'OK' button forcefully clicked via Actions API. (line: {sys._getframe().f_lineno})")

                            # 3. Synchronize thread layout: Wait for the SweetAlert backdrop container to leave the view entirely
                            # wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-container")))
                            # print("Success modal cleared. Main form view is ready for the next iteration.")
                            



                            print(f"Form population completed successfully. line: {sys._getframe().f_lineno}")
                            if i < total_buttons:
                                counter = counter + 1
                                continue

                        except Exception as e: # inner exception of option No, No, & photo upload
                            #print(f"An error occurred during automation: {e} line: {sys._getframe().f_lineno}")
                            print(f"An error occurred during form submission for card {card_number} line: {sys._getframe().f_lineno}")

                       


#####################################

                    except Exception as e:
                        print(f"Error clicking button #{i + 1} line: {sys._getframe().f_lineno}")
                        ok_button = driver.find_element(By.CSS_SELECTOR, "button.swal2-confirm")

                        # Force execution bypassing the UI layer
                        driver.execute_script("arguments[0].click();", ok_button)

                        # Note: If clicking a button causes a full page reload or changes the DOM structure, 
                        # you will need to re-fetch the element list inside the loop to avoid StaleElementReferenceException.
                        
                    except Exception as e:
                        print(f"Could not click button #{1}: line: {sys._getframe().f_lineno}")





                time.sleep(3)
                #Screening Logic will go here$$$$$$$$$$$$$$$$$$$$$$$$$$$$                

            except Exception as e: # exception of ration card for loop
                print(f"Pipeline crashed for card: {card_number} with error, Going in next Ration card line: {sys._getframe().f_lineno}")
                

    except Exception as e: # exception of login function
        print(f"An error occurred while filling the form: line: {sys._getframe().f_lineno}")




# start executing
login()