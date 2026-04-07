from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time, os
from PIL import Image
import pytesseract
import cv2

FY_type = ["PMAYG Cumulative progress till date", "2024-2025", "2025-2026"]
Scheme_type = ["PRADHAN MANTRI AWAAS YOJANA GRAMIN"]
State_type = ["CHHATTISGARH"]
District_type = ["SURAJPUR"]
block_type = ["BHAIYATHAN", "ODAGI", "PRATAPPUR", "PREMNAGAR", "RAMANUJNAGAR", "SURAJPUR"]

driver = webdriver.Chrome()
driver.get("https://dashboard.pmayg.dord.gov.in/netiay/masterlogin.aspx")
os.environ["OMP_THREAD_LIMIT"] = "1"

WAIT_SECONDS = 2
USER_ID = "CH26"
PASSWORD = "Dist@2026"
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
    captcha_input.send_keys(str(captcha_text))
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

# Login to the website
login()
mySleepFunction(WAIT_SECONDS)
driver.quit()
