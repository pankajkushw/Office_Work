import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time, os
from selenium.webdriver.common.action_chains import ActionChains



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
        ration_cards = ['226489928354', '226489993637']

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

                # Try a standard driver click first to allow Angular event bubbles to fire naturally
                search_btn.click()
                print("Standard browser search button click executed.")

                #Screening Logic will go here$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                # Find all active "Select" buttons in the table
                # This uses a partial text match or exact match on the text inside the button/link
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "custom-table")))

                # Target the buttons precisely using the class name 'action-btn' shown in your HTML
                select_buttons = driver.find_elements(By.CSS_SELECTOR, "button.action-btn")
                total_buttons = len(select_buttons)

                print(f"Found {total_buttons} matching 'Select' buttons.")
                counter = 0
                for i in range(total_buttons):

                    try:
                        #Second time date of visit, village has to be set and search again
                        counter = counter + 1
                        #if counter > 0:





                            
                            

                        # Re-fetch the elements inside the loop to ensure they are fresh
                        buttons = driver.find_elements(By.CSS_SELECTOR, "button.action-btn")
                        current_button = buttons[i]
                        
                        # Scroll the specific button into view before interacting
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", current_button)
                        time.sleep(0.5)
                        
                        print(f"Clicking Select button #{i + 1}...")
                        
                        # Use a reliable JavaScript click to bypass overlapping Angular components or overlay layers
                        driver.execute_script("arguments[0].click();", current_button)
                        
                            # Adjust this sleep duration based on how long your application takes to process a click

                        # Initialize the flag tracker at the start of each loop iteration
                        screening_already_done = False

                        try:
                            # 1. Target the explicit SweetAlert container that is open on your screen
                            print("Checking for active SweetAlert warning layout...")
                            swal_modal = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.CLASS_NAME, "swal2-modal"))
                            )
                            
                            # Verify the specific error message is inside the modal header
                            if "Screening Already Done" in swal_modal.text:
                                print("⚠️ Match found: 'Screening Already Done' alert verified.")
                                
                                # 2. Locate the precise SweetAlert confirm button using its dedicated library class
                                ok_btn = WebDriverWait(driver, 5).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm"))
                                )
                                
                                # 3. Force click via JavaScript to bypass any backdrop focus lock layers
                                driver.execute_script("arguments[0].click();", ok_btn)
                                print("SweetAlert 'OK' button successfully clicked.")
                                
                                # # 4. Wait for the SweetAlert dark backdrop container to leave the DOM hierarchy entirely
                                # WebDriverWait(driver, 5).until(
                                #     EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-container"))
                                # )
                                # 4. Wait for the SweetAlert dark backdrop container to leave the DOM hierarchy entirely
                                swal_container = driver.find_element(By.CLASS_NAME, "swal2-container")
                                WebDriverWait(driver, 5).until(EC.staleness_of(swal_container))

                                print("Modal faded out. Workspace cleared.")
                                
                                # Flip our loop bypass flag to True
                                screening_already_done = True

                        except Exception as e:
                            # If the modal doesn't exist, this block catch routes seamlessly into the normal flow
                            print(f"No active validation alert container intercepted: {e}")

                        # --- THE CRITICAL CONDITIONAL SKIP ---
                        if screening_already_done:
                            print("Skipping remaining form fields. Routing directly back to the next loop iteration...\n")
                            continue  # Breaks the current execution string and pulls the next record smoothly

                        # --- REST OF FORM SUBMISSION ROUTINE CONTINUES BELOW ---
                        print("Proceeding with normal report creation actions...")

                        time.sleep(2) 
                        #####################################
                        try:
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
                                is_visible = WebDriverWait(driver, 3).until(
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

                            # 3. Synchronize thread layout: Wait for the SweetAlert backdrop container to leave the view entirely
                            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-container")))
                            print("Success modal cleared. Main form view is ready for the next iteration.")




                            print("Form population completed successfully.")

                        except Exception as e: # inner exception of option No, No, & photo upload
                            print(f"An error occurred during automation: {e}")

                        finally:
                            # Keep the browser open for manual verification
                            input("Press Enter to close the browser...")
                            continue


#####################################

                    except Exception as e:
                        print(f"Error clicking button #{i + 1}: {e}")
                        
                        # Note: If clicking a button causes a full page reload or changes the DOM structure, 
                        # you will need to re-fetch the element list inside the loop to avoid StaleElementReferenceException.
                        
                    except Exception as e:
                        print(f"Could not click button #{1}: {e}")

                time.sleep(3)
                #Screening Logic will go here$$$$$$$$$$$$$$$$$$$$$$$$$$$$                

            except Exception as e: # exception of ration card for loop
                import traceback
                print(f"Pipeline crashed for card: {card_number}")
                print(traceback.format_exc())

    except Exception as e: # exception of login function
        print(f"An error occurred while filling the form: {e}")




# start executing
login()