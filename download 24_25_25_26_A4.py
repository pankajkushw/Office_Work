from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from PIL import Image
import pytesseract

FY_type = ["PMAYG Cumulative progress", "2024-2025", "2025-2026"]
Scheme_type = ["PRADHAN MANTRI AWAAS YOJANA GRAMIN"]
State_type = ["CHHATTISGARH"]
District_type = ["SURAJPUR"]
block_type = ["BHAIYATHAN", "ODAGI", "PRATAPPUR", "PREMNAGAR", "RAMANUJNAGAR", "SURAJPUR"]

driver = webdriver.Chrome()
driver.get("https://report.pmayg.dord.gov.in//netiay/PhysicalProgressReport/GapInProgressAccountVerificationCompletion.aspx")

def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)

def solveCaptcha():
    captcha_image = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_imgCaptcha")
    captcha_image.screenshot("captcha.png")
    captcha_text = pytesseract.image_to_string(Image.open("captcha.png"))
    print("Captcha Text:", captcha_text)
    result = eval(captcha_text)
    print("Captcha Result:", result)

    captcha_input = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtCaptcha")
    captcha_input.send_keys(str(result))
    time.sleep(1)
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
                # for block in block_type:
                #     dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlBlock"))
                #     dropdown.select_by_visible_text(block)
                #     print(block)
                #     print("Solving Captcha...")
                #     print( solveCaptcha())
                #     button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnSubmit")
                #     button.click()
                #     button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExport") 
                #     button.click()

                mySleepFunction(10) #Wait for captcha to be solved
                print("Data Downloaded for " + fy + " " + scheme + " " + state + " " + district)

driver.quit()