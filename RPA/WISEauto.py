from glob import glob
import time
import os.path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from pandas import read_excel
from json import loads

def exportReport(acc, fac, start, end):
    with open('accountconfigs.json', 'r') as f:
            data = loads(f.read())

    userDownloadPath = "C:\\Users\\" + os.getlogin() + "\\Downloads\\*.xlsx"
    downloadFolderBefore = glob(userDownloadPath)

    wisebots = read_excel('BNP Excel Sheet.xlsx', sheet_name='Facility list')
    facilityList = wisebots['FacilityName'].to_list()

    if fac in facilityList:
        index = facilityList.index(fac)
    else:
        raise Exception('No Wise Account for this Facility')

    #Removing pop-up for notification permissions
    chromeOptions = webdriver.ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications" : 2}
    chromeOptions.add_experimental_option("prefs", prefs)

    #Opening Wise website through chrome
    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chromeOptions)
    action = ActionChains(driver)
    driver.get(data['wiseDomain'])

    #Maximizing window to see all elements
    driver.maximize_window()

    #Inputting username and password
    interactor = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'username')))
    interactor.send_keys(wisebots['Account'][index])
    interactor = driver.find_element(By.NAME,"password")
    interactor.send_keys(wisebots['Password'][index])
    interactor = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="loginBtn"]/button')))
    interactor.click()

    #Selecting Report Center from Home Menu
    interactor = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '/html/body/div1/header/div[1]/div[3]/ul/li/a')))
    action.move_to_element(interactor).click().perform()
    try:
        interactor = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div1/header/div[1]/div[3]/ul/li/ul/li/div/div/div/ul/li[7]')))
    except:
        interactor = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '/html/body/div1/header/div[1]/div[3]/ul/li/a')))
        action.move_to_element(interactor).click().perform()
        interactor = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div1/header/div[1]/div[3]/ul/li/ul/li/div/div/div/ul/li[7]')))
    action.move_to_element(interactor).click().perform()

    #Seleccting Activity Report V2 from Billing
    interactor = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div1/div[2]/div[1]/ul/li[2]/a/span[1]')))
    action.move_to_element(interactor).click().perform()
    interactor = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div1/div[2]/div[1]/ul/li[2]/ul/li[2]/a/span')))
    action.move_to_element(interactor).click().perform()

    #Inputting Customer, Time From, Time To
    customer = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div[2]/form/div[1]/div[1]/organization-auto-complete/div/div')))
    action.click(customer).perform()
    wiseacc = acc[:10]
    time.sleep(2)
    for letter in acc:
        time.sleep(0.05)
        action.send_keys(letter).perform()

    time.sleep(4)
    action.send_keys(Keys.ENTER).perform()

    interactor = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div[2]/form/div[1]/div[3]/lt-date-time/div/input')))
    action.click(interactor).send_keys(start).perform()
    interactor = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div[2]/form/div[1]/div[4]/lt-date-time/div/input')))
    action.click(interactor).send_keys(end).perform()

    #Exporting
    interactor = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div[2]/form/div[2]/div[1]/input')
    action.move_to_element(interactor).click().perform()
    interactor = driver.find_element(By.XPATH, '//*[@id="sitecontent"]/div/div/div/div[2]/form/div[2]/div[2]/unis-waitting-btn/button')
    action.move_to_element(interactor).click().perform()

    timer = 0

    downloadWait = True
    while downloadWait:
        downloadFolderAfter = glob(userDownloadPath)
        if len(downloadFolderBefore) < len(downloadFolderAfter):
            downloadWait = False
        if timer == 15:
            try:
                action.click(customer).perform()
                time.sleep(1)
                for letter in acc:
                    time.sleep(0.05)
                    action.send_keys(letter).perform()

                time.sleep(4)
                action.send_keys(Keys.ENTER).perform()

                interactor = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div[2]/form/div[2]/div[1]/input')
                action.move_to_element(interactor).click().perform()
                interactor = driver.find_element(By.XPATH, '//*[@id="sitecontent"]/div/div/div/div[2]/form/div[2]/div[2]/unis-waitting-btn/button')
                action.move_to_element(interactor).click().perform()
            except:
                pass
        elif timer == 120:
            raise Exception("Could not locate account in WISE!")

        timer += 1
        time.sleep(1)

        try:
            interactor = driver.find_element(By.XPATH, '/html/body/div[7]/md-dialog/md-dialog-actions/button')
            interactor.click()
            raise Exception("No Data Found!")
        except:
            pass