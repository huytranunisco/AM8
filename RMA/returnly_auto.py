from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC


def launchBrowser():
    global driver

    #Navigating to the website
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get("https://dashboard.returnly.com/dashboard/users/login")

    return driver

def exportReport():
    userReturnly = 'FLAANT0001.rms@unisco.com'
    passReturnly = 'Syst0002'

    #Logging in to returnly
    interactor = driver.find_element(By.ID, 'user_email') 
    interactor.send_keys(userReturnly)
    interactor = driver.find_element(By.ID, 'user_password')
    interactor.send_keys(passReturnly)
    interactor = driver.find_element(By.XPATH, '//*[@id="new_user"]/div[3]/div/input')
    interactor.click()

    #Navigating to reports tab and exporting
    select = driver.find_element(By.XPATH, '//*[@id="sidebar"]/nav/ul/li[3]/a/span')
    select.click()
    select = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.NAME, 'UNIS RMA Report')))
    select.click()
    select = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, 'js-reporting-modal-submit')))
    select.click()


launchBrowser()
exportReport()
