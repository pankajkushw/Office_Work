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
target_village = "Karampur"
target_date = "10-22-2026" 
driver.implicitly_wait(10) 
def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)# Define an explicit wait timeout
wait = WebDriverWait(driver, 10)

# Configure the logging system
logging.basicConfig(
    filename='Karampur_29102026.log',         # Name of the log file
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

def for_single_entry(rc):
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
                print(f"[SKIP] Member '{member_name}' with Ration Card '{rc}' is {age} years old. Moving to next member.")
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
            "//div[contains(@class, 'calendar')]//*[text()='22'] | "
            "//span[text()='22'] | "
            "//*[normalize-space(text())='22']"
        )
        mySleepFunction(1)
        # Wait until the cell element is visible and ready to be tapped
        date_element = wait.until(EC.element_to_be_clickable((By.XPATH, date_xpath)))
        date_element.click()
        print("Successfully selected September 22th.")

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

        # continue page
        continue_btn_xpath = "//button[contains(., 'Continue')] | //span[contains(text(), 'Continue')]/ancestor::button"
        continue_btn = wait.until(EC.element_to_be_clickable((By.XPATH, continue_btn_xpath)))
        continue_btn.click()
        print(f"Successfully clicked 'Continue'. (line: {sys._getframe().f_lineno})")

        # Replace this list with your actual target Ration Card numbers
        ration_cards = ['226480249522', '226480280830', '226480290205', '226480390871', '226480534626', '226480616422', '226480691079', '226480747058', '226480812867', '226480821265', '226480885309', '226481004084', '226481071197', '226481135879', '226481162594', '226481220117', '226481336178', '226481343972', '226481467879', '226481549416', '226481558141', '226481582003', '226481676485', '226481720534', '226481876031', '226481925498', '226481982565', '226482058325', '226482138402', '226482162771', '226482290791', '226482339614', '226482348635', '226482367487', '226482385216', '226482395342', '226482410192', '226482485951', '226482677143', '226482724449', '226482775958', '226482830024', '226482852798', '226482859226', '226482860037', '226482997102', '226483145833', '226483152964', '226483403516', '226483467766', '226483621017', '226483700704', '226483796604', '226483911461', '226484121296', '226484296040', '226484372945', '226484451179', '226484454205', '226484484694', '226484545256', '226484555413', '226484556390', '226484633097', '226484683697', '226484821848', '226484826512', '226484847690', '226484964887', '226484980029', '226485024136', '226485256604', '226485275322', '226485322819', '226485494854', '226485530930', '226485689211', '226485833181', '226485838143', '226485914042', '226485975569', '226486074538', '226486085458', '226486172417', '226486410480', '226486541945', '226486545611', '226486611961', '226486632384', '226486836326', '226486840793', '226487022675', '226487044618', '226487119281', '226487211245', '226487231663', '226487473528', '226487488043', '226487523178', '226487560949', '226487563632', '226487575193', '226487592048', '226487729536', '226487745407', '226487813609', '226487938307', '226488029265', '226488133132', '226488174897', '226488195579', '226488219524', '226488230810', '226488526423', '226488531685', '226488653211', '226488682017', '226488722621', '226489063589', '226489116097', '226489160557', '226489209726', '226489266775', '226489290563', '226489404669', '226489486092', '226489522904', '226489619845', '226489645995', '226489804594', '226489938948', '226489946464', '223890581182', '223894092395', '223895965824', '226480005419', '226480012452', '226480025106', '226480055484', '226480074680', '226480101903', '226480105323', '226480174214', '226480193200', '226480209774', '226480287165', '226480330956', '226480332827', '226480361280', '226480400616', '226480445084', '226480447201', '226480604572', '226480607683', '226480637790', '226480664676', '226480696369', '226480767260', '226480767260', '226480770516', '226480880658', '226480934863', '226480951560', '226480961433', '226480969484', '226481004451', '226481045784', '226481049249', '226481056826', '226481090303', '226481156601', '226481169061', '226481188209', '226481200170', '226481238219', '226481239401', '226481245044', '226481290385', '226481303408', '226481356867', '226481360371', '226481457944', '226481467957', '226481490864', '226481493179', '226481514922', '226481536044', '226481559548', '226481559905', '226481572522', '226481613763', '226481627295', '226481640989', '226481640989', '226481674926', '226481683115', '226481706919', '226481709808', '226481712519', '226481737352', '226481749700', '226481764705', '226481773972', '226481816143', '226481820767', '226481842580', '226482023164', '226482030415', '226482116836', '226482123273', '226482126340', '226482205837', '226482220868', '226482229360', '226482249974', '226482271868', '226482312668', '226482354202', '226482410133', '226482410148', '226482429095', '226482442716', '226482446124', '226482489495', '226482492237', '226482532452', '226482532452', '226482592928', '226482595715', '226482599014', '226482602267', '226482617419', '226482645899', '226482731139', '226482768277', '226482779193', '226482781891', '226482838050', '226482842397', '226482884810', '226482908044', '226482921656', '226482929254', '226482978952', '226482994402', '226483003596', '226483022279', '226483033214', '226483058828', '226483142975', '226483183391', '226483201549', '226483232652', '226483261145', '226483279462', '226483298701', '226483300619', '226483366042', '226483417066', '226483446389', '226483453664', '226483467047', '226483474783', '226483477106', '226483604150', '226483639993', '226483664146', '226483667422', '226483700202', '226483713911', '226483718016', '226483755056', '226483775106', '226483776556', '226483806693', '226483915852', '226483921048', '226483952582', '226484010409', '226484057600', '226484057690', '226484061290', '226484063450', '226484071950', '226484086981', '226484088411', '226484148252', '226484175652', '226484179416', '226484180012', '226484182575', '226484188453', '226484209477', '226484218419', '226484227428', '226484229421', '226484249632', '226484277454', '226484299897', '226484313212', '226484333344', '226484369874', '226484370730', '226484379832', '226484385888', '226484429089', '226484431178', '226484431178', '226484439035', '226484496229', '226484514048', '226484559978', '226484592764', '226484596459', '226484598872', '226484600331', '226484603056', '226484614001', '226484619650', '226484670816', '226484687845', '226484801193', '226484805330', '226484837822', '226484839152', '226484874207', '226484876228', '226484893666', '226484947271', '226484955045', '226485010295', '226485013800', '226485018301', '226485028154', '226485035710', '226485039637', '226485207575', '226485219467', '226485247032', '226485261817', '226485279516', '226485330192', '226485430050', '226485463425', '226485497007', '226485501706', '226485513335', '226485519298', '226485546385', '226485580657', '226485584546', '226485602334', '226485633582', '226485645052', '226485661404', '226485753550', '226485791784', '226485826553', '226485874066', '226485880098', '226485901869', '226485930675', '226486034718', '226486081780', '226486085989', '226486108346', '226486114296', '226486116593', '226486142492', '226486160961', '226486177064', '226486194798', '226486197249', '226486203315', '226486244366', '226486270421', '226486275283', '226486332843', '226486350073', '226486358704', '226486383751', '226486390263', '226486390263', '226486403118', '226486403744', '226486404133', '226486424488', '226486477743', '226486487751', '226486503826', '226486521364', '226486530532', '226486542122', '226486556894', '226486584085', '226486584085', '226486673137', '226486682564', '226486696586', '226486714524', '226486745487', '226486769022', '226486788949', '226486834342', '226486840355', '226486959555', '226486974434', '226486984823', '226487023747', '226487039646', '226487074070', '226487076227', '226487095134', '226487109465', '226487121064', '226487161387', '226487176976', '226487190300', '226487193914', '226487272760', '226487276031', '226487301242', '226487325773', '226487353299', '226487376853', '226487471410', '226487453230', '226487540391', '226487564697', '226487571909', '226487591812', '226487632740', '226487655213', '226487705397', '226487707809', '226487775645', '226487804978', '226487846938', '226487849592', '226487878970', '226487919888', '226487923313', '226487932918', '226487976667', '226488066768', '226488100552', '226488192724', '226488206098', '226488212849', '226488267470', '226488278049', '226488363287', '226488466477', '226488522164', '226488524446', '226488530628', '226488531971', '226488536337', '226488553581', '226488571491', '226488599210', '226488605255', '226488627867', '226488638425', '226488667057', '226488683672', '226488686448', '226488704366', '226488757370', '226488796887', '226488907313', '226488983851', '226489036405', '226489055671', '226489069015', '226489090524', '226489176840', '226489180457', '226489197725', '226489218964', '226489253499', '226489254753', '226489417513', '226489529321', '226489539898', '226489561982', '226489564289', '226489584256', '226489586112', '226489646226', '226489679133', '226489702521', '226489736524', '226489776416', '226489795683', '226489826479', '226489878451', '226489903241', '226489917157', '226489984089', '226481031582', '226481086311', '226481246697', '226481287369', '226481295091', '226481329177', '226481370727', '226481445235', '226481475437', '226481490056', '226481498137', '226481538910', '226481568664', '226481731382', '226481767777', '226481783742', '226481816740', '226481914283', '226481931836', '226482100875', '226482205405', '226482214647', '226482219572', '226482331544', '226482465109', '226482473179', '226482486773', '226482597161', '226482597425', '226482648129', '226482662478', '226482714442', '226482748625', '226482834104', '226482918139', '226483056531', '226483111689', '226483187660', '226483197585', '226483346894', '226483360580', '226483400789', '226483585926', '226483646792', '226483652305', '226483661577', '226483783307', '226483976844', '226484272146', '226484319895', '226484355015', '226484394012', '226484530730', '226484597586', '226484643774', '226484663169', '226484776610', '226484842055', '226485316537', '226485593171', '226485593536', '226485632613', '226485706648', '226485732124', '226485747657', '226485821857', '226485851903', '226485879336', '226485991440', '226486053393', '226486077050', '226486162149', '226486247960', '226486260561', '226486261393', '226486397072', '226486515964', '226486535743', '226486580336', '226486632946', '226486668463', '226486799132', '226486849531', '226486853603', '226486873055', '226486938586', '226486943358', '226487015741', '226487080693', '226487102828', '226487197251', '226487202125', '226487253122', '226487256814', '226487388396', '226487447259', '226487591416', '226487626974', '226487650645', '226487666863', '226487678954', '226487742085', '226487762811', '226487805756', '226487864158', '226487893332', '226487990250', '226488039906', '226488053587', '226488085912', '226488093473', '226488106469', '226488122234', '226488123151', '226488266071', '226488311893', '226488425215', '226488549478', '226488613216', '226488617835']
        
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
                    for_single_entry(card_number)
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