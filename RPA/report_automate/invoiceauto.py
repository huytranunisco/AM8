import BNPauto
import WISEauto
from datetime import date, timedelta
import os.path
from os import rename
from pandas import read_excel
from calendar import monthrange
from shutil import copy, move
from glob import glob
from time import time

start_time = time()

def getInvoice(acc, facility, startP, endP, accName, cycle, wise = False):
    try:
        downloaddir = 'C:\\Users\\' + os.getlogin() + '\\Downloads'
        accsdir = 'C:\\Users\\' + os.getlogin() +'\\Desktop\\Discrepancy Reports\\Accounts'

        if wise:
            WISEauto.exportReport(acc, facility, startP, endP)
            newName = accName + '-' + facility + '-' + cycle + '-Activity_Report' + weekNum + '.xlsx'
            copyPath = 'C:\\Users\\' + os.getlogin() + '\\Desktop\\Discrepancy Reports\\Accounts\\03 - Current Activity reports\\' + newName
        else:
            facility, invoiceName = BNPauto.exportHandle(acc, facility, startP, endP, accName)
            if facility == False or invoiceName == False: 
                return False

            newName = accName + '-' + facility + '-' + cycle + '-Invoice' + weekNum + '.xlsx'
            copyPath = 'C:\\Users\\' + os.getlogin() + '\\Desktop\\Discrepancy Reports\\Accounts\\02 - Current Invoices\\' + newName + '.xlsx'
        
        fileList = list(filter(os.path.isfile, glob(downloaddir + '\\*.xlsx')))
        fileList.sort(key=lambda x: os.path.getmtime(x))
        file = fileList[len(fileList) - 1]
        fileEnd = os.path.basename(file)
        
        if wise:
            move(os.path.join(downloaddir, fileEnd), os.path.join(accsdir, '01 - Historical Activity reports', fileEnd))
            copy(os.path.join(accsdir, '01 - Historical Activity reports', fileEnd), copyPath)
        else:
            move(os.path.join(downloaddir, fileEnd), os.path.join(accsdir, '00 - Historical Invoices', fileEnd))
            copy(os.path.join(accsdir, '00 - Historical Invoices', fileEnd), copyPath)

            if fileEnd == 'Invoice[Multi].xlsx':
                rename(os.path.join(accsdir, '00 - Historical Invoices', fileEnd), os.path.join(accsdir, '00 - Historical Invoices', accName + "-" + facility + "-" + fileEnd))

        print(f'Downloaded and Copied {newName} \n')

        return True

    except Exception as e:
        if hasattr(e, 'message'):
            print(e.message)
            invoiceAccs['Downloaded'][index] = e.message
            invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
        else:
            print('An error occured at ', e.args, e.__doc__)
            invoiceAccs['Downloaded'][index] = e.args, e.__doc__
            invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)

def getBimonthlyDate(todayDate):
    if todayDate.day < 16:
        previousMonth = todayDate.month - 1 if todayDate.month != 1 else 12
        previousYear = todayDate.year - 1 if todayDate.month == 12 else todayDate.year
        start = date(previousYear, previousMonth, 16)
        end = date(previousYear, previousMonth, monthrange(previousYear, previousMonth)[1])
    elif today.day >= 16:
        start = date(todayDate.year, todayDate.month, 1)
        end = date(todayDate.year, todayDate.month, 15)
    
    return start, end

def getWeeklyDate(todayDate):
    idx = (todayDate.weekday() + 1) % 7 # MON = 0, SUN = 6 -> SUN = 0 .. SAT = 6

    sun = today - timedelta(7+idx)
    sat = today - timedelta(7+idx-6)
    
    return sun, sat

invoiceAccs = read_excel('BNP Excel Sheet.xlsx', sheet_name='Account_Fac_Freq')
bimonthly = False
weekly = False

today = date.today()
dayName = today.strftime("%A")

if dayName == 'Sunday' and (today.day == 16 or today.day == 1):
    bimonthly = True
    weekly = True
elif dayName == 'Sunday':
    invoiceAccs = invoiceAccs[invoiceAccs['BillingFreq'] == 'Weekly']
    weekly = True
elif today.day == 16 or today.day == 1:
    invoiceAccs = invoiceAccs[invoiceAccs['BillingFreq'] == 'Bimonthly']
    bimonthly = True
else:
    raise Exception("Not a valid day to run program!")


for index in invoiceAccs.index:
    fac = invoiceAccs['Facility Name'][index]
    bnpName = invoiceAccs['BNP Account Name'][index]
    wiseName = invoiceAccs['Wise Account Name'][index]
    accName = invoiceAccs['AccountName'][index]

    print(f'Getting Invoice for {accName} ({fac})...')
    if bimonthly:
        pass
    elif bimonthly and weekly:
        if invoiceAccs['BillingFreq'][index] == 'Bimonthly':
            #Assuming is ran on 1st and 16th
            startBimonthly, endBimonthly = getBimonthlyDate(today)

            bnpStart = startBimonthly.strftime("%m/%d/%y")
            bnpEnd = endBimonthly.strftime("%m/%d/%y")
        
            wiseStart = startBimonthly.strftime("%y-%m-%d")
            wiseEnd = endBimonthly.strftime("%y-%m-%d")

            bnpInvoice = getInvoice(bnpName, fac, bnpStart, bnpEnd, accName, 'Bimonthly')
            wiseInvoice = False

            if (bnpInvoice):
                wiseInvoice = getInvoice(wiseName, fac, wiseStart, wiseEnd, accName, 'Bimonthly', True)
                if (wiseInvoice):
                    invoiceAccs['Downloaded'][index] = True
                    invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
                else:
                    invoiceAccs['Downloaded'][index] = 'No Wise Invoice'
                    invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
                    continue
            else:
                invoiceAccs['Downloaded'][index] = 'No BNP Invoice'
                invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
                continue

        elif invoiceAccs['BillingFreq'][index] == 'Weekly':

            startWeekly, endWeekly = getWeeklyDate(today)

            bnpStart = startWeekly.strftime("%m/%d/%y")
            bnpEnd = endWeekly.strftime("%m/%d/%y")
            wiseStart = startWeekly.strftime("%y-%m-%d")
            wiseEnd = endWeekly.strftime("%y-%m-%d")

            bnpInvoice = getInvoice(bnpName, fac, bnpStart, bnpEnd, accName, 'Weekly', False)
            wiseInvoice = False
                
            if (bnpInvoice):
                wiseInvoice = getInvoice(wiseName, fac, wiseStart, wiseEnd, accName, 'Weekly', True)
                if(wiseInvoice):
                    invoiceAccs['Downloaded'][index] = True
                    invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
                else:
                    invoiceAccs['Downloaded'][index] = 'No Wise Invoice'
                    invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
                    continue
            else:
                invoiceAccs['Downloaded'][index] = 'No BNP Invoice'
                invoiceAccs.to_excel('AccountsDone.xlsx', sheet_name='Account_Fac_Freq', index=False)
                continue

print("Invoices Downloaded!")
print("--- %s seconds ---" % (time() - start_time))
x = input("Press Enter to finish. ")