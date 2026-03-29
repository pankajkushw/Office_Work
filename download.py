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

driver = webdriver.Chrome()
driver.get("https://report.pmayg.dord.gov.in//netiay/PhysicalProgressReport/GapInProgressAccountVerificationCompletion.aspx")

captcha_image = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_imgCaptcha")
captcha_image.screenshot("captcha.png")
captcha_text = pytesseract.image_to_string(Image.open("captcha.png"))
print("Captcha Text:", captcha_text)
result = eval(captcha_text)
print("Captcha Result:", result)

try:
    captcha_input = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtCaptcha")
    captcha_input.send_keys(str(result))
    time.sleep(5)

except:
    print("Failed to input captcha text.")