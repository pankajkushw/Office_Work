from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time, os, io
from PIL import Image
import pytesseract
import cv2
import pandas as pd

Sanction_fy_year = ["As per Generated Financial Year"]
FY_type = ["2026-2027"]
Scheme_type = ["PRADHAN MANTRI AWAAS YOJANA GRAMIN"]
State_type = ["CHHATTISGARH"]
District_type = ["SURAJPUR"]
#block_type = ["BHAIYATHAN", "ODAGI", "PRATAPPUR", "PREMNAGAR", "RAMANUJNAGAR", "SURAJPUR"]
block_type = ["PREMNAGAR"]
FTO_type = ["Total No. of FTO Generated"]

driver = webdriver.Chrome()
driver.get("https://report.pmayg.dord.gov.in/netiay/EFMSReport/SNASparsh_FtoTransactionSummaryReport.aspx")
os.environ["OMP_THREAD_LIMIT"] = "1"

WAIT_SECONDS = 3
driver.implicitly_wait(10) 
def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)

def solveCaptcha():

    captcha_image = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_imgCaptcha")
    captcha_image.screenshot("captcha.png")    
    captcha_text = pytesseract.image_to_string("captcha.png", config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789+-*/()')
    mySleepFunction(2) #Wait for captcha to be solved
    print("Captcha Text:", captcha_text)
    result = eval(captcha_text)
    
    print("Captcha Result:", result)
    
    captcha_input = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtCaptcha")
    captcha_input.send_keys(str(result))
    return result    

for sfy in Sanction_fy_year:
    dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlGenSan"))
    dropdown.select_by_visible_text(sfy)
    # 1. Change dropdown
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
                for district in District_type: 
                    dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlDistrict"))
                    dropdown.select_by_visible_text(district)
                    print(district)
                    print("Solving Captcha...")
                    #print( solveCaptcha())
                    for block in block_type:
                        dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlBlock"))
                        dropdown.select_by_visible_text(block)
                        print(block)
                        for ftype in FTO_type:
                            dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddl_detail"))
                            dropdown.select_by_visible_text(ftype)
                            print(ftype)
                            
                            button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
                            button.click()

                            # link_elements = driver.find_elements(By.PARTIAL_LINK_TEXT, "FTO")
                            # urls = [link.get_attribute("href") for link in link_elements if link.get_attribute("href")]

                            # print(f"Found {len(urls)} links to traverse.")
                            # all_tables_data = []

                            # # Save the main window handle to return to the master report page
                            # main_window = driver.current_window_handle
                            # all_tables_data = []

                            # 1. Target the links inside the "FTO File Name" column explicitly
                            # Save the main report window handle context
                            main_window = driver.current_window_handle
                            all_tables_data = []

                            print("Waiting for the master FTO transaction summary table rows to populate...")

                            try:
                                # 1. Broadly wait for the data table container to be visible on screen
                                WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.XPATH, "//table[contains(@id, 'grd')] | //table[contains(@id, 'gv')]"))
                                )
                                time.sleep(2) # Vital buffer for dynamic ASP.NET grid rows to fully populate
                                
                                # 2. A more flexible XPath that targets any link text inside the rows of your report table
                                fto_links_xpath = "//table[contains(@id, 'grd') or contains(@id, 'gv')]//tr[td]//td/a"
                                
                                # Extract the live clickable list elements
                                links_elements = driver.find_elements(By.XPATH, fto_links_xpath)
                                total_fto_files = len(links_elements)
                                
                                print(f"Found {total_fto_files} FTO file links to process.")

                            except Exception as e:
                                print(f"Error locating report data grid table links: {e}")
                                total_fto_files = 0

                            if total_fto_files > 0:
                                for index in range(total_fto_files):
                                    print(f"\nProcessing file [{index + 1}/{total_fto_files}]")
                                    
                                    try:
                                        # Force focus back to the master list table view 
                                        driver.switch_to.window(main_window)
                                        
                                        # Re-fetch links live to completely prevent StaleElementReferenceException
                                        links = driver.find_elements(By.XPATH, fto_links_xpath)
                                        target_link = links[index]
                                        file_name = target_link.text.strip()
                                        print(f"Clicking on FTO File: {file_name}")
                                        
                                        # Click the link using JavaScript execution for 100% reliable tracking
                                        driver.execute_script("arguments[0].click();", target_link)
                                        time.sleep(2) # Safe buffer for the browser to load the data view or tab

                                        # 2. Check if a new tab window was spawned by the click action
                                        all_windows = driver.window_handles
                                        if len(all_windows) > 1:
                                            driver.switch_to.window(all_windows[-1])

                                        # 3. Wait for the detail table container to load completely
                                        WebDriverWait(driver, 15).until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, "table#ctl00_ContentPlaceHolder1_grdDetail"))
                                        )
                                        time.sleep(1)
                                        
                                        # 4. Extract the outerHTML structure
                                        target_table_element = driver.find_element(By.CSS_SELECTOR, "table#ctl00_ContentPlaceHolder1_grdDetail")
                                        table_html = target_table_element.get_attribute("outerHTML")
                                                                
                                        # 5. Load the target data into your Pandas collection arrays
                                        dfs = pd.read_html(io.StringIO(table_html))
                                        if dfs:
                                            df = dfs[0]
                                            # --- NORMALIZE MULTI-INDEX HEADERS IF PRESENT ---
                                            # If Pandas reads an extra structural banner row, flatten it out
                                            if isinstance(df.columns, pd.MultiIndex):
                                                df.columns = df.columns.get_level_values(-1)

                                            df = df.iloc[:-1]
                                            
                                            # Map tracking metadata tags
                                            df["FTO_File_Name"] = file_name
                                            df["Source_Index"] = index + 1
                                            
                                            all_tables_data.append(df)
                                            print(f"Successfully scraped {len(df)} rows from {file_name}.")

                                    except Exception as e:
                                        print(f"Could not parse data for item index {index + 1}: {e}")

                                    # 6. CRITICAL CLEANUP: If a new tab window popped open, close it to return to the master table
                                    if len(driver.window_handles) > 1:
                                        driver.close()
                                        driver.switch_to.window(main_window)
                                    else:
                                        # If it updated on the SAME page, use browser back to return to the master view
                                        # Uncomment the line below ONLY if the details view opens on the same tab without creating a new window
                                        # driver.back()
                                        pass

                                # 7. Compile everything and save
                                if all_tables_data:
                                    combined_df = pd.concat(all_tables_data, ignore_index=True)
                                    combined_df.to_excel("scraped_fto_details_output.xlsx", index=False)
                                    print("\nSaved all collected FTO datasets to Excel file successfully!")

                    mySleepFunction(5) #Wait for captcha to be solved
                    print("Data Downloaded for " + fy + " " + scheme + " " + state + " " + district)
driver.quit()



