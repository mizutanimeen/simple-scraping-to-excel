from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome import service
import os
from dotenv import load_dotenv
import time
import openpyxl

def main(driver):
    wait = WebDriverWait(driver, 10)
    driver.get(os.getenv("URL"))
    wait.until(EC.presence_of_all_elements_located)
    time.sleep(1)
    elements = driver.find_elements("tag name", os.getenv("TAG_NAME"))

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    row = 1
    for e in elements:
        if e.text != "":
            sheet.cell(row=row, column=1, value=e.text)
            row += 1
            time.sleep(1)
    filename = 'output.xlsx'
    workbook.save(filename)
    if os.getenv("IS_OPEN") == "True":
        os.system(f'start excel.exe {filename}')
    
if __name__ == "__main__":
    load_dotenv()
    service = service.Service(executable_path=os.getenv("DRIVER_PATH"))
    driver = webdriver.Chrome(service=service)
    try:
        main(driver)
    finally:
        driver.quit()