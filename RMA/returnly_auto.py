import email
from email import policy
from email.parser import BytesParser
import pandas as pd
import imaplib
import getpass
from bs4 import BeautifulSoup
import urllib.request
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

def exportReport():
    print('Enter password for Returnly!')
    userReturnly = 'FLAANT0001.rms@unisco.com'
    passReturnly = getpass.getpass()

    #Logging in to returnly
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.get("https://dashboard.returnly.com/dashboard/users/login")
        interactor = driver.find_element(By.ID, 'user_email') 
        interactor.send_keys(userReturnly)
        interactor = driver.find_element(By.ID, 'user_password')
        interactor.send_keys(passReturnly)
        print('Attempting to login...')
        select =  WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="new_user"]/div[3]/div/input')))
        select.click()
    except Exception as e:        
        print('Error: ', e)
        driver.quit()

    #Navigating to reports tab and exporting
    try:
        select = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="sidebar"]/nav/ul/li[3]/a/span')))
        select.click()
        select = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, 'UNIS RMA Report')))
        select.click()
        select = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, 'js-reporting-modal-submit')))
        select.click()
        print('Download succesful, report has been sent to email!')
        driver.quit()
    except Exception as e:
        print('Failed to export report\n','Error: ', e)
        driver.quit()
        md

def downloadReport():
    domain = 'webmail.unisco.com'
    userEmail = 'FLAANT0001.rms@unisco.com'
    userPass = getpass.getpass()
    folder = input('Please enter location you want to download report (remove ""): ')

    try:
        M = imaplib.IMAP4_SSL(domain)
        M.login(userEmail, userPass)
        
        M.SELECT('INBOX')
        rv, data = M.search(None, 'Unseen')
        if rv == False:
            print("No report found yet, please wait!")
            pass
        
        for num in data[0].split():
            rv, data = M.fetch(num, '(RFC822)')
            if rv == False:
                print("ERROR getting message"), num
                return
            print("Writing message "), num
            f = open('%s/%s.eml' %(folder, num), 'wb')
            f.write(data[0][1])
            f.close()

        file = input('Please enter report location (remove ""): ')

        with open(file) as email_file:
            email_message = email.message_from_file(email_file)
        if email_message.is_multipart():
            for part in email_message.walk():
                message = str(part.get_payload(decode=True))
                print(message)

    except Exception as e:
        print('Error: ', e)

def formatReport():

    file = input('Please enter location of report (remove ""): ')
    
    df = pd.read_csv(file)

    df = df[['RMA', 'Original Order', 'Tracking Number', 'Shopper Email', 'Shipped From Name', 'Shipped From Address 1', 'Shipped From City', 'Shipped From State',
             'Shipped From Zip', 'Shipped From Country Code', 'Barcode', 'Return Created Date', 'Return Delivered Date']]

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

    df = df[['ClientID', 'RMA', 'Reference', 'Reverse Type', 'Return Party', 'Return From Address 1', 'Return From City', 'Return From State', 'Return From Postal Code', 'Return From Country',
                'Phone', 'Email', 'RMA Expiration Date', 'Ship Method', 'Shipment Carrier', 'Pro / Tracking Number', 'BOL', 'ETA', 'Note', 'Item Name', 'UPC', 'Buyer ID', 'Return Qty', 'UOM', 'Serial Number', 'Return Created Date',
                'Return Delivered Date']]

    df.to_excel('./RMA/Transformed.xlsx', index=False)

if __name__ == '__main__':
    downloadReport()

