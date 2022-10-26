import pandas as pd
import imaplib
import os
import time
import email
from datetime import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

def exportReport():

    #Logging in to returnly
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get('https://dashboard.returnly.com/dashboard/users/login')
    interactor = driver.find_element(By.XPATH, '//*[@id="user_email"]') 
    interactor.send_keys(userReturnly)
    interactor = driver.find_element(By.XPATH, '//*[@id="user_password"]')
    interactor.send_keys(passReturnly)
    select =  WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="new_user"]/div[3]/div/input')))
    select.click()

    #Navigating to reports tab and exporting
    driver.get('https://dashboard.returnly.com/dashboard/reports')
    select = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="reconciliation-cards-container"]/article[4]/div/nav/ul/li[2]/span')))
    select.click()
    select = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="js-reporting-modal-submit"]')))
    select.click()
    driver.quit()

#Connecting to email domain through iMAP
def loginEmail():
    domain = 'webmail.unisco.com'
    M = imaplib.IMAP4_SSL(domain)
    M.login(userEmail, passEmail)

    return M

#Searching for specific email
def searchEmail():
    M.select('Inbox')
    status, data = M.search(None, '(UNSEEN FROM "help@returnly.com" SUBJECT "Your Returnly report is ready")')
    newest_data = getNewestEmail(data)

    #If no new email is found it searches repetitively
    try:
        if newest_data[0].isdigit() == False:
            pass
        while newest_data[0].isdigit() == False:
            newest_data = searchEmail()
            if newest_data[0].isdigit():
                break
    except IndexError:
        pass

    if status != 'OK':
        print('Error occured while searching... ', e)

    return newest_data

#Retrieving the newest email
def getNewestEmail(newest_data):
    try:
        ids = newest_data[0]
        id_list = ids.split()
        latest_emails = id_list
        keys = map(int, latest_emails)
        news_keys = sorted(keys, reverse=True)
        str_keys = [str(e) for e in news_keys]
        return str_keys
    except IndexError:
        pass

#Downloading the email that was retrieved
def downloadReport():
    #Searching for an unread email from returnly if not found an index error arrises and it goes back into searching
    search = False
    while (search == False):
        try:
            time.sleep(5)
            data = searchEmail()
            search = bool(searchEmail())
        except IndexError:
            continue

    #Writing the data retrieved from the email to the eml file
    for num in data[0].split():
        status2, data2 = M.fetch(num, '(RFC822)')
        if status2 != 'OK':
            print("Error downloading the email...")
            return
        numy = str(num) + '.eml'
        x = os.path.join(f"C:\\Users\\{user}\\Downloads\\", numy)
        f = open(x, 'wb')
        f.write(data2[0][1])
        f.close()
    numy = str(data[0]) + '.eml'
    y = os.path.join(f"C:\\Users\\{user}\\Downloads\\", numy)

    #Iterating through the messages in the email and retrieving the download link
    messageList = []
    with open(y) as email_file:
        email_message = email.message_from_file(email_file)
    if email_message.is_multipart():
        for part in email_message.walk():
            message = str(part.get_payload(decode=True))
            messageList.append(message)

    #Eliminating all other text besides the first link (the download link)
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
        
    #Logging into returnly again to retrieve report from the link
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get('https://dashboard.returnly.com/dashboard/users/login')
    interactor = driver.find_element(By.XPATH, '//*[@id="user_email"]') 
    interactor.send_keys(userReturnly)
    interactor = driver.find_element(By.XPATH, '//*[@id="user_password"]')
    interactor.send_keys(passReturnly)
    select =  WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="new_user"]/div[3]/div/input')))
    select.click()
    driver.get(linksList[0])
    download_wait(30)
    driver.quit()

#Checks download folder for file and waits a certain amount of time, if time is exceeded it times out
def download_wait(timeout):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(f"C:\\Users\\{user}\\Downloads\\")
        for fname in files:
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds

#Formatting the report to liking
def formatReport():
    files = [os.path.join(f'C:\\Users\\{user}\\Downloads\\', x) for x in os.listdir(f'C:\\Users\\{user}\\Downloads\\') if x.endswith(".csv")]
    newest = max(files , key = os.path.getctime) #Retrieves the latest .csv file in downloads folder

    file = newest
    
    df = pd.read_csv(file)

    df = df[['RMA Number', 'Original Order ID', 'Return Tracking Number', 'Customer Email', 'Shipped From Name', 'Shipped From Address 1', 'Shipped From City', 'Shipped From State',
             'Shipped From ZIP', 'Shipped From Country Code', 'Barcode', 'Return Initiated Date', 'Return Delivered Date']]

    newColDict = {'ClientID':'FLAANT0001', 'Reverse Type':'Consumer Return', 'Phone':'', 'RMA Expiration Date':'', 'Ship Method':'Small Parcel', 'Shipment Carrier':'USPS',
                  'BOL':'', 'ETA':'', 'Note':'', 'UPC':'', 'Buyer ID':'', 'Return Qty':1, 'UOM':'EA', 'Serial Number':''}

    newColKeys = newColDict.keys()

    for index, colName in enumerate(newColKeys):
        df[colName] = newColDict[colName]

    #Renaming columns to the template format
    df = df.rename(columns={'RMA Number' : 'RMA', 'Original Order ID' : 'Reference', 'Shipped From Name' : 'Return Party', 'Shipped From Address 1' : 'Return From Address 1',
                            'Shipped From City' : 'Return From City', 'Shipped From State' : 'Return From State', 'Shipped From ZIP' : 'Return From Postal Code', 
                            'Shipped From Country Code' : 'Return From Country', 'Customer Email' : 'Email', 'Return Tracking Number' : 'Pro / Tracking Number',
                            'Barcode' : 'Item Name', 'Return Initiated Date' : 'Return Created Date'})

    #Reordering columns to template format
    df = df[['ClientID', 'RMA', 'Reference', 'Reverse Type', 'Return Party', 'Return From Address 1', 'Return From City', 'Return From State', 'Return From Postal Code', 'Return From Country',
                'Phone', 'Email', 'RMA Expiration Date', 'Ship Method', 'Shipment Carrier', 'Pro / Tracking Number', 'BOL', 'ETA', 'Note', 'Item Name', 'UPC', 'Buyer ID', 'Return Qty', 'UOM', 'Serial Number', 'Return Created Date',
                'Return Delivered Date']]

    df['UPC'] = df['Item Name']

    time = datetime.now()
    filename = time.strftime("%m-%d-%Y %H-%M-%S")
    df.to_excel(f"C:\\Users\\{user}\\Downloads\\Transformed_{filename}_PDT.xlsx", index = False)

if __name__ == '__main__':
    user = os.getlogin()
    userReturnly = 'FLAANT0001.rms@unisco.com'
    passReturnly = 'Syst0002'
    userEmail = 'FLAANT0001.rms@unisco.com'
    passEmail = 'Syst0001'

    try:
        print('Exporting report...')
        exportReport()
        print('\nDownloading report...')
        M = loginEmail()
        downloadReport()
        print('\nFormatting report...')
        formatReport()
        print('\nProcess completed!\n')

    except Exception as e:
        print('Erorr: ', e)
