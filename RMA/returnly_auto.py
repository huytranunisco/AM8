from argparse import Namespace
from datetime import datetime, timedelta
import pandas as pd
import imaplib
import email
import os
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
    userEmail = ''
    passEmail = ''

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
    

    server = 'outlook.office365.com'
    user = 'gilbert.castellanos@unisco.com'
    password = 'bgr388d'
    outputdir = '/Downloads'
    subject = 'Data Exports' #subject line of the emails you want to download attachments from

    # connects to email client through IMAP
    def connect(server, user, password):
        m = imaplib.IMAP4_SSL(server)
        m.login(user, password)
        m.select()
        return m

    # downloads attachment for an email id, which is a unique identifier for an
    # email, which is obtained through the msg object in imaplib, see below 
    # subjectQuery function. 'emailid' is a variable in the msg object class.

    def downloaAttachmentsInEmail(m, emailid, outputdir):
        resp, data = m.fetch(emailid, "(BODY.PEEK[])")
        email_body = data[0][1]
        mail = email.message_from_bytes(email_body)
        if mail.get_content_maintype() != 'multipart':
            return
        for part in mail.walk():
            if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
                open(outputdir + '/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))

    # download attachments from all emails with a specified subject line
    # as touched upon above, a search query is executed with a subject filter,
    # a list of msg objects are returned in msgs, and then looped through to 
    # obtain the emailid variable, which is then passed through to the above 
    # downloadAttachmentsinEmail function

    def subjectQuery(subject):
        m = connect(server, user, password)
        m.select("Inbox")
        typ, msgs = m.search(None, '(SUBJECT "' + subject + '")')
        msgs = msgs[0].split()
        for emailid in msgs:
            downloaAttachmentsInEmail(m, emailid, outputdir)

    subjectQuery(subject)
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
