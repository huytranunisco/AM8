import email
import imaplib
import os
import webbrowser
import pandas as pd
from email.header import decode_header
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

def exportReport():
    userReturnly = 'FLAANT0001.rms@unisco.com'
    passReturnly = 'Syst0002'

    #Logging in to returnly
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.get("https://dashboard.returnly.com/dashboard/users/login")
        interactor = driver.find_element(By.ID, 'user_email') 
        interactor.send_keys(userReturnly)
        interactor = driver.find_element(By.ID, 'user_password')
        interactor.send_keys(passReturnly)
        select =  WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="new_user"]/div[3]/div/input')))
        select.click()
        print('Returnly login succesful!')
    except Exception as e:        
        print('Login failed\n','Error: ', e)
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
        print('Download succesful, report has been sent to email!')
        driver.quit()
    except Exception as e:
        print('Failed to export report\n','Error: ', e)
        driver.quit()

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

def downloadReport():
    userEmail = 'FLAANT0001.rms@unisco.com'
    userPass = 'Syst0001'

    try:
        imap_server = "webmail.unisco.com"
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(userEmail, userPass)
        print('Email login successful!')

        status, messages = imap.select("INBOX")
        N = 1
        messages = int(messages[0])

        for i in range(messages, messages-N, -1):
            # fetch the email message by ID
            res, msg = imap.fetch(str(i), "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    # parse a bytes email into a message object
                    msg = email.message_from_bytes(response[1])
                    # decode the email subject
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        # if it's a bytes, decode to str
                        subject = subject.decode(encoding)
                    # decode email sender
                    From, encoding = decode_header(msg.get("From"))[0]
                    if isinstance(From, bytes):
                        From = From.decode(encoding)
                    print("Subject:", subject)
                    print("From:", From)
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        print(body)
                    if content_type == "text/html":
                        # if it's HTML, create a new HTML file and open it in browser
                        folder_name = clean(subject)
                        if not os.path.isdir(folder_name):
                            # make a folder for this email (named after the subject)
                            os.mkdir(folder_name)
                        filename = "index.html"
                        filepath = os.path.join(folder_name, filename)
                        # write the file
                        open(filepath, "w").write(body)
                        # open in the default browser
                        webbrowser.open(filepath)
                    print("="*100)
                    """
                    # if the email message is multipart
                    if msg.is_multipart():
                        # iterate over email parts
                        for part in msg.walk():
                            # extract content type of email
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))
                            try:
                                # get the email body
                                body = part.get_payload(decode=True).decode()
                            except:
                                pass
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                # print text/plain emails and skip attachments
                                print(body)
                            elif "attachment" in content_disposition:
                                # download attachment
                                filename = part.get_filename()
                                if filename:
                                    folder_name = clean(subject)
                                    if not os.path.isdir(folder_name):
                                        # make a folder for this email (named after the subject)
                                        os.mkdir(folder_name)
                                    filepath = os.path.join(folder_name, filename)
                                    # download attachment and save it
                                    open(filepath, "wb").write(part.get_payload(decode=True))
                    else:
                        # extract content type of email
                        content_type = msg.get_content_type()
                        # get the email body
                        body = msg.get_payload(decode=True).decode()
                        if content_type == "text/plain":
                            # print only text email parts
                            print(body)
                    if content_type == "text/html":
                        # if it's HTML, create a new HTML file and open it in browser
                        folder_name = clean(subject)
                        if not os.path.isdir(folder_name):
                            # make a folder for this email (named after the subject)
                            os.mkdir(folder_name)
                        filename = "index.html"
                        filepath = os.path.join(folder_name, filename)
                        # write the file
                        open(filepath, "w").write(body)
                        # open in the default browser
                        webbrowser.open(filepath)
                    print("="*100)"""
        # close the connection and logout
        imap.close()
        imap.logout()
    except Exception as e:
        print('Login failed\n','Error: ', e)

    """
    print('\nNow attempting to download the report from email...')
    #Login to email service
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.get("https://webmail.unisco.com/")
        interactor = driver.find_element(By.XPATH, '//*[@id="User"]')
        interactor.send_keys(userEmail)
        interactor = driver.find_element(By.XPATH, '//*[@id="Password"]')
        interactor.send_keys(userPass)
        select =  WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Logon"]')))
        select.click()
        print('Login succesful!')
    except Exception as e:
        print('Failed to login to email\n','Error: ', e)
        driver.quit()

    #Navigating email to download report
    try:
        select = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@title="Your Returnly report is ready"]')))
        select.click()
    except Exception as e:
        print('Failed to navigate to email\n','Error: ', e)
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

if __name__ == '__main__':
    exportReport()
    downloadReport()

