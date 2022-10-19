import email
from email import policy
from email.parser import BytesParser
import pandas as pd
import imaplib
import getpass
import os
import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

def exportReport():
    global userReturnly
    global passReturnly

    print('Enter password for Returnly!')
    userReturnly = 'FLAANT0001.rms@unisco.com'
    passReturnly = 'Syst0002' #getpass.getpass()

    #Logging in to returnly
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.get('https://dashboard.returnly.com/dashboard/users/login')
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

def downloadReport():
    domain = 'webmail.unisco.com'
    userEmail = 'FLAANT0001.rms@unisco.com'
    userPass = 'Syst0001' #getpass.getpass()
    folder = input('Please enter the folder location you want to download the report to (remove ""): ')

    while os.path.exists(folder) != True:
        folder = input('Please enter a valid folder location...: ')

    try:
        M = imaplib.IMAP4_SSL(domain)
        M.login(userEmail, userPass)
        
        while True:
            M.SELECT('INBOX')
            rv, data = M.search(None, 'Unseen')
            print(rv, data)
            if rv != 'OK':
                print("No report found yet... please wait...")
                time.sleep(10)
                rv, data = M.search(None, 'Unseen')
                if rv == 'OK':
                    print('Report has been found... continuing...')
                    break
            else:
                print('Report has been found... continuing...')
                break       

        for num in data[0].split():
            while True:
                rv, data = M.fetch(num, '(RFC822)')
                if rv != 'OK':
                    print("Error downloading the email... Please make sure that the email is 'unread' and it is in the inbox!"), num
                    quit()
                else:
                    break
            print('Downloading the email...'), num
            f = open('%s/%s.eml' %(folder, num), 'wb')
            f.write(data[0][1])
            f.close()

        file = input('Please enter report location with file name at end (remove "", filepath/b"12".eml): ')

        messageList = []
        with open(file) as email_file:
            email_message = email.message_from_file(email_file)
        if email_message.is_multipart():
            for part in email_message.walk():
                message = str(part.get_payload(decode=True))
                messageList.append(message)

        a = messageList[1]
        newText = ""
        urlFlag = 0
        linksList = []

        for i in range(len(a)):
            if(a[i]=='<'):
                urlFlag = 1
                continue
            if(a[i]=='>'):
                urlFlag = 0
                linksList.append(newText)
                newText = ''
                continue
            if(urlFlag==1):
                newText = newText + a[i]

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.get('https://dashboard.returnly.com/dashboard/users/login')
        interactor = driver.find_element(By.ID, 'user_email') 
        interactor.send_keys('FLAANT0001.rms@unisco.com')
        interactor = driver.find_element(By.ID, 'user_password')
        interactor.send_keys('Syst0002')
        print('Attempting to login...')
        select =  WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="new_user"]/div[3]/div/input')))
        select.click()
        print('Succesfully logged in!')
        driver.get(linksList[0])
        time.sleep(5)
        driver.quit()

    except Exception as e:
        print('Error: ', e)

def formatReport():

    file = input('Please enter location of the downloaded report (remove ""): ')
    
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

