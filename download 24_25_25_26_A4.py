from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time, os
from PIL import Image
import pytesseract
import cv2

FY_type = ["PMAYG Cumulative progress", "2024-2025", "2025-2026"]
Scheme_type = ["PRADHAN MANTRI AWAAS YOJANA GRAMIN"]
State_type = ["CHHATTISGARH"]
District_type = ["SURAJPUR"]
block_type = ["BHAIYATHAN", "ODAGI", "PRATAPPUR", "PREMNAGAR", "RAMANUJNAGAR", "SURAJPUR"]

driver = webdriver.Chrome()
driver.get("https://report.pmayg.dord.gov.in//netiay/PhysicalProgressReport/GapInProgressAccountVerificationCompletion.aspx")
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

    # image = cv2.imread("captcha.png")
    # gray = cv2.imread(image, cv2.COLOR_BGR2GRAY)
    # thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    # custom_config = r'--oem 3 --psm 7'

    
    captcha_text = pytesseract.image_to_string("captcha.png", config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789+-*/()')
    mySleepFunction(2) #Wait for captcha to be solved
    print("Captcha Text:", captcha_text)
    result = eval(captcha_text)
    
    print("Captcha Result:", result)
    
    captcha_input = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtCaptcha")
    captcha_input.send_keys(str(result))
    return result    

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
            if(state == "CHHATTISGARH"):
                #Get State Data 
                print("Solving Captcha...")
                print( solveCaptcha())
                button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
                button.click()
                
                button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExport")
                button.click()
                
            
            for district in District_type: 
                dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlDistrict"))
                dropdown.select_by_visible_text(district)
                print(district)
                print("Solving Captcha...")
                print( solveCaptcha())
                button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
                button.click()
                
                button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExport") 
                button.click()
                
                #download block data for
                for block in block_type:
                    dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlBlock"))
                    dropdown.select_by_visible_text(block)
                    print(block)
                    print("Solving Captcha...")
                    print( solveCaptcha())
                    button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
                    button.click()
                    
                    button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExport") 
                    button.click()
                    

                mySleepFunction(5) #Wait for captcha to be solved
                print("Data Downloaded for " + fy + " " + scheme + " " + state + " " + district)

driver.quit()