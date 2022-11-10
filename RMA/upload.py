import os
import json
import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

def Upload(file):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    actions = ActionChains(driver)
    driver.get('https://stage.logisticsteam.com/#/login')
    driver.maximize_window()

    element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div1/div/div[2]/div/div/div/div/form/div/input[1]')))
    element.send_keys(userWISE)

    element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div1/div/div[2]/div/div/div/div/form/div/input[2]')))
    element.send_keys(passWISE)

    element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="loginBtn"]/button')))
    actions.move_to_element(element).click().perform()

    time.sleep(3)
    driver.get('https://stage.logisticsteam.com/#/rms/returnManagement/rma')
    driver.maximize_window()

    driver.switch_to().frame(1)

    element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.LINK_TEXT, 'Import RMA')))
    actions.move_to_element(element).click().perform()

    #element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, ""xpath=//a[@onclick='ToImportRMA()']")))
    #actions.move_to_element(element).click().perform()

    #uploadelement = webdriverwait(driver, 20).until(ec.element_to_be_clickable((by.xpath, "(xpath=//input[@id='files']"))));
    #uploadelement.send_keys('c:\\users\\gcastellanos\\downloads\\transformed_11-03-2022 10-31-52_pdt.xlsx')
    
    
if __name__ == '__main__':
    try:
        user = os.getlogin()
        with open(f'C:\\Users\\gcastellanos\\Documents\\GitHub\\AM8\\RMA\\config.json', 'r') as f:
            data = json.loads(f.read())
        userWISE = data['userWISE']
        passWISE = data['passWISE']
        Upload('a')
    except Exception as e:
        print(e)