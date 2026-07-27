from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time, os
from PIL import Image
import pytesseract
import cv2
import pandas as pd

Sanction_fy_year = ["As per Generated Financial Year"]
FY_type = ["2026-2027"]
Scheme_type = ["PRADHAN MANTRI AWAAS YOJANA GRAMIN"]
State_type = ["CHHATTISGARH"]
District_type = ["SURAJPUR"]
block_type = ["BHAIYATHAN", "ODAGI", "PRATAPPUR", "PREMNAGAR", "RAMANUJNAGAR", "SURAJPUR"]
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
                    for ftype in FTO_type:
                        dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddl_detail"))
                        dropdown.select_by_visible_text(ftype)
                        print(ftype)
                        
                        button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
                        button.click()

                        link_elements = driver.find_elements(By.PARTIAL_LINK_TEXT, "FTO")
                        urls = [link.get_attribute("href") for link in link_elements if link.get_attribute("href")]

                        print(f"Found {len(urls)} links to traverse.")
                        all_tables_data = []
                        # 4. Loop through each extracted URL
                        # 2. Traverse each link
                        for index, url in enumerate(urls):
                            
                            print(f"Processing page [{index + 1}/{len(urls)}]: {url}")
                            driver.get(url)
                            
                            # Wait for the data table on the sub-page to completely render
                            # Replace 'table.target-data-table' with your actual target table selector
                            WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_gvdetails"))
                            )
                            time.sleep(1) # Safe buffer for AJAX data rendering
                            
                            # 3. Extract the inner HTML text of the table
                            target_table_element = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_gvdetails")
                            table_html = target_table_element.get_attribute("outerHTML")
                                                       
                            # 4. Use Pandas to convert HTML directly into a DataFrame
                            # read_html returns a list of dataframes; grab the first one [0]
                            try:
                                dfs = pd.read_html(table_html)
                                if dfs:
                                    df = dfs[0]
                                    
                                    # Optional: Add a metadata column tracking which URL this row came from
                                    df["Source_URL"] = url 
                                    
                                    all_tables_data.append(df)
                                    print(f"Successfully scraped {len(df)} rows.")
                            except Exception as e:
                                print(f"Could not parse table on page {index + 1}: {e}")

                        # 5. Combine and export the compiled dataset
                        if all_tables_data:
                            combined_df = pd.concat(all_tables_data, ignore_index=True)
                            
                            # Export Option A: To Excel Workbook
                            excel_filename = "scraped_tables_output.xlsx"
                            combined_df.to_excel(excel_filename, index=False)
                            print(f"\nSaved all data to Excel: {excel_filename}")
                            
                            # Export Option B: To CSV File (Uncomment if preferred)
                            # csv_filename = "scraped_tables_output.csv"
                            # combined_df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
                            # print(f"\nSaved all data to CSV: {csv_filename}")
                        else:
                            print("No table data was found or collected.")


                    mySleepFunction(5) #Wait for captcha to be solved
                    print("Data Downloaded for " + fy + " " + scheme + " " + state + " " + district)

driver.quit()



