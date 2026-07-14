import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time, os
from pathlib import Path
from PIL import Image
import pytesseract
import cv2
import requests
import threading

w_FY_type = ["2016-2017", "2017-2018", "2018-2019", "2019-2020", "2020-2021", "2022-2023", "2024-2025", "2025-2026"]
FY_type = ["2024-2025", "2025-2026", "PMAYG Cumulative progress till date" ]
w_Scheme_type = ["PMAYG"]
Scheme_type = ["PRADHAN MANTRI AWAAS YOJANA GRAMIN"]
State_type = ["CHHATTISGARH"]
District_type = ["SURAJPUR"]
block_type = ["BHAIYATHAN", "ODAGI", "PRATAPPUR", "PREMNAGAR", "RAMANUJNAGAR", "SURAJPUR"]
data_type = ["Sanctioned Year"]
Panchayat_type = ["All"]
Category_type = ["All"]
Progress_type = ["All"]

Service = Service(timeout = 300)
driver = webdriver.Chrome(service=Service)
driver.get("https://dashboard.pmayg.dord.gov.in/netiay/masterlogin.aspx")
os.environ["OMP_THREAD_LIMIT"] = "1"
driver.page_load_strategy = 'eager'

WAIT_SECONDS = 2
USER_ID = "CH26"
PASSWORD = "Dspr@202627"
FY = "2025-2026"    
driver.implicitly_wait(10) 
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
    alert = driver.find_element(By.NAME, "btnClose")
    if alert:
        alert.click()
        print("Alert closed.")


    #calling to download A2 Report1649	1540	1330
    A2_Report()
    Aawas_Plus_Report()
    #WorkinProgress_Report()

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

def WorkinProgress_Report():
    w_FY_type = ["2016-2017", "2017-2018", "2018-2019", "2019-2020", "2020-2021", "2022-2023", "2024-2025", "2025-2026"]
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
                                
                                mySleepFunction(3)
                                print("Data Downloaded for " + fy + " " + scheme + " " + fy + " " + block + " " + panchayat + " " + category_type + " " + progress_type)
                                new_file_name = f"WorkinProgress_{block}_{fy}"
                                rename_latest_download(new_file_name)
                                mySleepFunction(1)



def rename_latest_download(new_filename):
    # 1. Automatically locate the default user Downloads folder
    download_dir = Path.home() / "Downloads"
    
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
    new_file_path = download_dir / f"{new_filename}{file_extension}"
    
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
