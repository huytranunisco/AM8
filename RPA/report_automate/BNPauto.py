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
from difflib import SequenceMatcher
from pandas import read_excel
import facilities
import os.path

def facilityMatcher(givenF):
    highestratio = 0
    facName = ''
    facs = facilities.facilityList
    while (True):
        for f in facs:
            ratio = SequenceMatcher(None, f.lower(), givenF.lower()).ratio()
            if ratio > highestratio:
                highestratio = ratio
                facName = f
            if ratio == 1:
                return facName
        return facName

def exportHandle(acc, fac, start, end, accName):
    billTo = acc
    facility = fac
    facility = facilityMatcher(facility)
    periodStart = start
    periodEnd = end
    userbnp = 'wiserpa'
    pwbnp = '#rpa#1234'

    #Getting BNP
    chromeOptions = webdriver.ChromeOptions()
    prefs = {'safebrowsing.disable_download_protection' : True}
    chromeOptions.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chromeOptions)
    action = ActionChains(driver)
    driver.get("http://bnp.unisco.com/")

    #Logging into BNP
    interactor = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'inputUserName')))
    interactor.send_keys(userbnp)
    interactor = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'inputPassword')))
    interactor.send_keys(pwbnp)
    interactor = driver.find_element(By.XPATH,"/html/body/div/footer/div/button")
    interactor.click()

    #Clicking Sales Module from Module Dropdown Menu
    select = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'TS_span_menu')))
    select.click()
    select = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="headmenu_mn_active"]/div/ul/li[1]')))
    select.click()

    #Clicking Invoice Management from Invoice Dropdown Menu
    select = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="headmenu"]/li[3]/span')))
    select.click()
    select = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/header/div[1]/ul/li[3]/div/ul/li[1]/a')))
    select.click()

    #Inputting information on the Invoice Management Page and Searching

    #Bill To
    interactor = driver.find_element(By.XPATH, '//*[@id="sitecontent"]/div[2]/div[1]/div[6]/span/span/input')
    interactor.send_keys(billTo)
    time.sleep(0.5)
    interactor.send_keys(Keys.ENTER)

    #Facility
    select = Select(driver.find_element(By.ID, 'ddlFacility'))
    select.select_by_visible_text(facility)

    #Invoice Status
    interactor = driver.find_element(By.XPATH, '//*[@id="sitecontent"]/div[2]/div[2]/div[5]/div/div/input')
    interactor.send_keys('Check')
    time.sleep(0.5)
    interactor.send_keys(Keys.ENTER)

    #Period Start
    interactor = driver.find_element(By.ID, 'inputPeriodStart')
    interactor.send_keys(periodStart)

    #Period End
    interactor = driver.find_element(By.ID, 'inputPeriodEnd')
    interactor.send_keys(periodEnd)

    #Category
    interactor = driver.find_element(By.XPATH, '//*[@id="sitecontent"]/div[2]/div[2]/div[6]/div/input')
    interactor.send_keys('Handling')

    #Clicking Search
    interactor = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[3]/div[11]/button')
    interactor.click()
    interactor = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[3]/div[11]/button')
    interactor.click()
    time.sleep(2)

    table = driver.find_element(By.XPATH, '//*[@id="invoicegrid"]/div[3]/table')
    rows = table.find_elements(By.TAG_NAME, 'tr')
    if len(rows) == 0:
        raise Exception(f'No invoice found for {accName}_{facility}!')
    elif len(rows) > 1:
        for index in range(len(rows)):
            xpath = '//*[@id=\"invoicegrid\"]/div[3]/table/tbody/tr[' + str(index + 1) + ']/td[1]/label'
            interactor = driver.find_element(By.XPATH, xpath)
            action.move_to_element(interactor).perform()
            interactor.click()
        
        invoiceName = 'Multi'
    else:
        row = rows[0]
        col = row.find_elements(By.TAG_NAME, 'td')[3]
        invoiceName = col.text

        #Checking first invoice
        interactor = driver.find_element(By.XPATH, '//*[@id=\"invoicegrid\"]/div[3]/table/tbody/tr[1]/td[1]/label')
        action.move_to_element(interactor).perform()
        interactor.click()

    #Exporting Handling Invoice
    time.sleep(2)
    interactor = driver.find_element(By.ID, 'btnExportInvoiceDetail')
    action.move_to_element(interactor).perform()
    interactor.click()

    user = os.getlogin()

    path = 'C:\\Users\\' + user + '\\Downloads\\Invoice[' + invoiceName + '].xlsx'

    while not os.path.exists(path):
        time.sleep(1)
        if os.path.isfile(path):
            break
        
    return facility, path

def invoiceToReport(acc, fac, billingPeriod, invoiceNum):
    reportName = acc + '-' + fac + '-' + billingPeriod + '.xlsx'

    path = 'Invoice[' + invoiceNum + '].xlsx'
    report = read_excel(path, sheet_name='Item Summary')
    
    new_cols = ['Category', 'InvoiceNumber', 'Header Billing Period Start', 'Header Billing Period End', 'ItemName', 'Description', 'UnitPrice', 'Qty']
    report = report.reindex(columns=new_cols)

    report.rename(columns={'Qty' : 'BNP Qty'}, inplace=True)
    report['WISE Qty'] = ''
    report['CSR Qty'] = ''

    report.loc[report['Category'] != 'Outbound', 'Category'] = ''
    report.loc[report['ItemName'] == 'HANDLING OUT', 'Category'] = 'Outbound'
    report.loc[report['ItemName'] == 'HANDLING IN', 'Category'] = 'Inbound'


    report.to_excel(reportName, index=False)
    
    print("Discrepancy Report has been made!")

    return reportName