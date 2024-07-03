from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import numpy as np
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
import pandas as pd
from openpyxl import Workbook

PATH = "C:\Program Files (x86)\chromedriver-win64\chromedriver.exe"

cService = webdriver.ChromeService(executable_path=PATH)   # instantiate
options = webdriver.ChromeOptions()  # instantiate      
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service = cService, options=options)

driver.get("https://www.google.com/maps/place/Adidas+Outlet+-+DVOM/@5.244397,100.29339,12z/data=!4m12!1m2!2m1!1sadidas!3m8!1s0x304ab9dfc5f8f8c9:0xbe30133fcec8059c!8m2!3d5.244397!4d100.4375856!9m1!1b1!15sCgZhZGlkYXMiA4gBAZIBEHNwb3J0c3dlYXJfc3RvcmXgAQA!16s%2Fg%2F11c2kzbw8q?entry=ttu")


def MineReviews():
    action = ActionChains(driver)
    allReviews = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[class="hh2c6 G7m0Af"]')))
    allReviews = driver.find_element(By.CSS_SELECTOR, '[class="hh2c6 G7m0Af"]')
    allReviews.click()
    
    comments = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[class = "wiI7pd"]')))
    comments = driver.find_elements(By.CSS_SELECTOR, '[class = "wiI7pd"]')

    while len(comments)<1000:
        var = len(comments)
        scroll_origin = ScrollOrigin.from_element(comments[len(comments)-1])
        action.scroll_from_origin(scroll_origin, 0, 1000).perform()
        time.sleep(2)
        comments = driver.find_elements(By.CSS_SELECTOR, '[class = "wiI7pd"]')

        if len(comments) == var:
            le+=1
            if le > 20:
                break
        else:
            le = 0

    for line in comments:
        print(line.text)
        comment.append(line.text)

    np.savetxt('review.csv',comment, delimiter=';' ,fmt='%s', encoding="utf-8")
    df = pd.read_csv("review.csv", on_bad_lines='skip')
    df.to_excel('review.xlsx')


def main():
    global wait 
    global comment
    comment = []
    wait = WebDriverWait(driver, 10)

    MineReviews()

main()



        