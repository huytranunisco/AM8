import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from pandas import read_excel

def exportReport(acc, fac, start, end):
    periodStart = start
    periodEnd = end
    userbnp = 'wiserpa'
    pwbnp = '#rpa#1234'

    wisebots = read_excel('BNP Excel Sheet.xlsx', sheet_name='Facility list')
    facilityList = wisebots['FacilityName'].to_list()

    if fac in facilityList:
        index = facilityList.index(fac)

    chromeOptions = webdriver.ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications" : 2}
    chromeOptions.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chromeOptions)
    action = ActionChains(driver)
    driver.get("https://wise.logisticsteam.com/v2/#/login")

    driver.maximize_window()

    interactor = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div1/div/div[2]/div/div/div/form/input[1]')))
    interactor.send_keys(wisebots['Account'][index])
    interactor = driver.find_element(By.NAME,"password")
    interactor.send_keys(wisebots['Password'][index])
    interactor = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="loginBtn"]/button')))
    interactor.click()

    interactor = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div1/header/div[1]/div[3]/ul/li/a')))
    action.move_to_element(interactor).click().perform()
    interactor = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div1/header/div[1]/div[3]/ul/li/ul/li/div/div/div/ul/li[7]')))
    action.move_to_element(interactor).click().perform()

    interactor = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div1/div[2]/div[1]/ul/li[2]/a')))
    action.move_to_element(interactor).click().perform()
    interactor = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div1/div[2]/div[1]/ul/li[2]/ul/li[2]/a')))
    action.move_to_element(interactor).click().perform()

    interactor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div[2]/form/div[1]/div[1]/organization-auto-complete/div/div')))
    action.click(interactor).send_keys(acc).perform()
    time.sleep(3)
    action.move_to_element_with_offset(interactor, 0, 10).click().perform()
    interactor = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div[2]/form/div[1]/div[3]/lt-date-time/div/input')))
    action.click(interactor).send_keys(start).perform()
    interactor = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div[2]/form/div[1]/div[4]/lt-date-time/div/input')))
    action.click(interactor).send_keys(end).perform()

exportReport('HIH Logistics, Inc.', 'Seabrook', '2022-10-01', '2022-10-15')