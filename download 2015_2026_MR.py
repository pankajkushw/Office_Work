from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time, os
from PIL import Image
import pytesseract
import cv2

def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)


FY_file_link = {"2015-16":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2015-2016&source=national&rbl1=0&rbl2=0&Digest=3gCuxmKvAMiJqCM9xdhuLQ",
                "2016-17":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2016-2017&source=national&rbl1=0&rbl2=0&Digest=ou40gSt7W1zGl5/7fAj7Yg",
                "2017-18":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2017-2018&source=national&rbl1=0&rbl2=0&Digest=QCg7N01gOK4vR/AzPgoTYQ",
                "2018-19":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2018-2019&source=national&rbl1=0&rbl2=0&Digest=E85qzeq83DlFatwBwOdfWw",
                "2019-20":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2019-2020&source=national&rbl1=0&rbl2=0&Digest=t01puYwSkjB42xGDe2I7nQ",
                "2020-21":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2020-2021&source=national&rbl1=0&rbl2=0&Digest=TY6ZiDTTlK9SxJhBmE1Aag",
                "2021-22":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2021-2022&source=national&rbl1=0&rbl2=0&Digest=pFCIY2TBHEwW9/UMaB6gJw",
                "2022-23":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2022-2023&source=national&rbl1=0&rbl2=0&Digest=Eu45/d65A4m5hwDXKyIzAg",
                "2023-24":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2023-2024&source=national&rbl1=0&rbl2=0&Digest=Ali2tMZbVMdxWQAdzmgprw",
                "2024-25":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2024-2025&source=national&rbl1=0&rbl2=0&Digest=kKbwLC4M57hU55YFddS6ug",
                "2025-26":"https://mnregaweb4.nic.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2025-2026&source=national&rbl1=0&rbl2=0&Digest=7/cekcr/Th+9BV08XjdumQ",
                "2026-27":"https://mnregaweb4.dord.gov.in/netnrega/ongo_comp_pds_wrk_rpt_new.aspx?page=D&short_name=&state_name=CHHATTISGARH&state_code=33&district_name=SURAJPUR&district_code=3326&fin_year=2026-2027&source=national&rbl1=0&rbl2=0&Digest=GHaoif+8dHG7sw0ZLznChA"
}

driver = webdriver.Chrome()
os.environ["OMP_THREAD_LIMIT"] = "1"
for link in FY_file_link:
    print(f"Downloading MR Report for FY:  {link}")
    driver.get(FY_file_link[link])
    download_click = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_LinkButton1")
    download_click.click()
    mySleepFunction(1) #Wait for captcha to be solved
driver.quit()    
