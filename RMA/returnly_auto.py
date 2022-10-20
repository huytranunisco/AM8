from asyncio.windows_events import NULL
import email
from email import policy
from email.parser import BytesParser
from pickle import EMPTY_LIST
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
        print('Login succesful... continuing')
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

    a = exportReport.has_been_called = True
    return a
a = exportReport.has_been_called = False

def downloadReport():
    domain = 'webmail.unisco.com'
    userEmail = 'FLAANT0001.rms@unisco.com'
    userPass = 'Syst0001' #getpass.getpass()

    try:
        M = imaplib.IMAP4_SSL(domain)

        M.login(userEmail, userPass)

        M.select('Inbox')

        status, data = M.search(None, '(UNSEEN FROM "help@returnly.com")')

        if data != [s for s in data if s.isdigit()]:
            print("No report found yet... please wait...")
            while True:
                status, data = M.search(None, '(UNSEEN FROM "help@returnly.com" SUBJECT "Your Returnly report is ready")')
                if data == [s for s in data if s.isdigit()]:
                    print('Report has been found... continuing...')
                    break

        for num in data[0].split():
            while True:
                status, data = M.fetch(num, '(RFC822)')
                if status != 'OK':
                    print("Error downloading the email... Please make sure that the report email is at the top and 'unread'!"), num
                    quit()
                else:
                    break
            print('Downloading the email...'), num
            numy = str(num) + '.eml'
            x = os.path.join(r"C:\Users\gcastellanos\Downloads", numy)
            f = open(numy, "x")
            f = open(x)
            f.write(data[0][1])
            f.close()
        
        file = input('Please enter file path of report (remove ""): ')
        while os.path.exists(file) != True:
            file = input('Please enter a valid report location...: ')

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
        print('File downloading...')
        time.sleep(5)
        driver.quit()

    except Exception as e:
        print('Error: ', e)

    b = downloadReport.has_been_called = True
    return b

b = downloadReport.has_been_called = False

def formatReport():
    file = input('Please enter file path of the downloaded report (remove ""): ')
    while os.path.exists(file) != True:
        file = input('Please enter a valid file...: ')
    
    df = pd.read_csv(file)

    df = df[['RMA Number', 'Original Order ID', 'Return Tracking Number', 'Customer Email', 'Shipped From Name', 'Shipped From Address 1', 'Shipped From City', 'Shipped From State',
             'Shipped From ZIP', 'Shipped From Country Code', 'Barcode', 'Return Initiated Date', 'Return Delivered Date']]

    newColDict = {'ClientID':'FLAANT0001', 'Reverse Type':'Consumer Return', 'Phone':'', 'RMA Expiration Date':'', 'Ship Method':'Small Parcel', 'Shipment Carrier':'USPS',
                  'BOL':'', 'ETA':'', 'Note':'', 'UPC':'', 'Buyer ID':'', 'Return Qty':1, 'UOM':'EA', 'Serial Number':''}

    newColKeys = newColDict.keys()

    for index, colName in enumerate(newColKeys):
        df[colName] = newColDict[colName]

    df = df.rename(columns={'RMA Number' : 'RMA', 'Original Order ID' : 'Reference', 'Shipped From Name' : 'Return Party', 'Shipped From Address 1' : 'Return From Address 1',
                            'Shipped From City' : 'Return From City', 'Shipped From State' : 'Return From State', 'Shipped From ZIP' : 'Return From Postal Code', 
                            'Shipped From Country Code' : 'Return From Country', 'Customer Email' : 'Email', 'Return Tracking Number' : 'Pro / Tracking Number',
                            'Barcode' : 'Item Name', 'Return Initiated Date' : 'Return Created Date'})

    df = df[['ClientID', 'RMA', 'Reference', 'Reverse Type', 'Return Party', 'Return From Address 1', 'Return From City', 'Return From State', 'Return From Postal Code', 'Return From Country',
                'Phone', 'Email', 'RMA Expiration Date', 'Ship Method', 'Shipment Carrier', 'Pro / Tracking Number', 'BOL', 'ETA', 'Note', 'Item Name', 'UPC', 'Buyer ID', 'Return Qty', 'UOM', 'Serial Number', 'Return Created Date',
                'Return Delivered Date']]

    df['UPC'] = df['Item Name']

    df.to_excel('./RMA/Transformed.xlsx', index=False)
    print('Transformation complete!')

    c = formatReport.has_been_called = True
    return c
c = formatReport.has_been_called = False

if __name__ == '__main__':
    #exportReport()
    downloadReport()
    #formatReport()

    if a and b and c:
        print('Process fully completed with no issues!')
    elif a == False:
        print('Something went wrong when attempting to export report...')
    elif b == False:
        print('Something went wrong when attempting to download report...')
    elif c == False:
        print('Somethingw went wrong when attempting to format report...')