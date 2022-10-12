import pandas as pd
import imaplib
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC


def launchBrowser():
    global driver

    #Navigating to the website
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get("https://dashboard.returnly.com/dashboard/users/login")

    return driver

def exportReport():
    userReturnly = 'xxx'
    passReturnly = 'xxx'

    #Logging in to returnly
    try:
        interactor = driver.find_element(By.ID, 'user_email') 
        interactor.send_keys(userReturnly)
        interactor = driver.find_element(By.ID, 'user_password')
        interactor.send_keys(passReturnly)
        select =  WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="new_user"]/div[3]/div/input')))
        select.click()
        print('Login succesful!')
    except Exception as e:        
        print('Login failed\n', 'Error: ', e)
        driver.quit()

    #Navigating to reports tab and exporting
    try:
        select = driver.find_element(By.XPATH, '//*[@id="sidebar"]/nav/ul/li[3]/a/span')
        select.click()
        select = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, 'UNIS RMA Report')))
        select.click()
        select = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, 'js-reporting-modal-submit')))
        print('Download succesful report has been sent to email!')
        driver.quit()
    except Exception as e:
        print('Failed to export report\n', 'Error: ', e)
        driver.quit()

def retrieveReport():
    domain = 'unisco.com'
    userEmail = 'xxx'
    userPass = 'xxx'

    imap_server = 'imap-mail.outlook.com'
    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(userEmail, userPass)

"""
    #Logging in to email
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.get("https://login.microsoftonline.com/")
        select = driver.find_element(By.C, '//*[@id="lightboxTemplateContainer"]/div[2]/div[1]/div/div/div/div/div[2]/div/div/div/div/div[2]') 
        select.click()
        interactor = driver.find_element(By.XPATH, '//*[@id="credentialList"]/div[3]/div[1]/div/div[2]')
        interactor.send_keys(domain)
    except Exception as e:
        print('Failed to login\n', 'Error: ', e)
"""

def formatReport():

    file = "C:\\Users\\gcastellanos\\Desktop\\Flag_&_Anthem_Returnly_Return_Details_created_at_2022-09-20_2022-09-20_PDT.csv"

    df = pd.read_csv(file)

    df = df[['RMA', 'Original Order', 'Tracking Number', 'Shopper Email', 'Shipped From Name', 'Shipped From Address 1', 'Shipped From City', 'Shipped From State',
             'Shipped From Zip', 'Shipped From Country Code', 'Barcode']]

    newColDict = {'ClientID':'FLAANT0001', 'Reverse Type':'Consumer Return', 'Phone':'', 'RMA Expiration Date':'', 'Ship Method':'Small Parcel', 'Shipment Carrier':'USPS',
                  'BOL':'', 'ETA':'', 'Note':'', 'UPC':'', 'Buyer ID':'', 'Return Qty':1, 'UOM':'EA', 'Serial Number':''}

    newColKeys = newColDict.keys()

    for index, colName in enumerate(newColKeys):
        df[colName] = newColDict[colName]

    df = df.rename(columns={'Original Order' : 'Reference', 'Shipped From Name' : 'Return Party', 'Shipped From Address 1' : 'Return From Address 1',
                            'Shipped From City' : 'Return From City', 'Shipped From State' : 'Return From State', 'Shipped From Zip' : 'Return From Postal Code', 
                            'Shipped From Country Code' : 'Return From Country', 'Shopper Email' : 'Email', 'Tracking Number' : 'Pro / Tracking Number', 
                            'Barcode' : 'Item Name'})
                    
    df['Reference'] = df['Reference'].str.replace('#', '')

    df.to_excel('Output.xlsx', index=False)

launchBrowser()
exportReport()
retrieveReport()
