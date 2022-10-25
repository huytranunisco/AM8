import BNPauto
import WISEauto
from datetime import date, timedelta
import os.path
from pandas import read_excel
from calendar import monthrange
from shutil import copy, move
from glob import glob
from time import time

start_time = time()

def getInvoice(acc, facility, startP, endP, accName, cycle, wise = False, weekNum = ''):
    try:
        downloaddir = 'C:\\Users\\' + os.getlogin() + '\\Downloads'
        accsdir = 'C:\\Users\\' + os.getlogin() +'\\Desktop\\Discrepancy Reports\\Accounts'

        if wise:
            WISEauto.exportReport(acc, facility, startP, endP)
            newName = accName + '-' + facility + '-' + cycle + '-Activity_Report' + weekNum + '.xlsx'
            newPath = os.path.join(accsdir, '01 - Historical Activity reports', newName)
            copyPath = 'C:\\Users\\' + os.getlogin() + '\\Desktop\\Discrepancy Reports\\Accounts\\03 - Current Activity reports\\' + newName
        else:
            facility, invoiceName = BNPauto.exportHandle(acc, facility, startP, endP, accName)
            if facility == False or invoiceName == False: 
                return False

            newName = accName + '-' + facility + '-' + cycle + '-Invoice' + weekNum + '.xlsx'
            newPath = os.path.join(accsdir, '00 - Historical Invoices', newName)
            copyPath = 'C:\\Users\\' + os.getlogin() + '\\Desktop\\Discrepancy Reports\\Accounts\\02 - Current Invoices\\' + newName + '.xlsx'
        
        fileList = list(filter(os.path.isfile, glob(downloaddir + '\\*.xlsx')))
        fileList.sort(key=lambda x: os.path.getmtime(x))
        file = fileList[len(fileList) - 1]
        fileEnd = os.path.basename(file)
        
        if wise:
            move(os.path.join(downloaddir, fileEnd), os.path.join(accsdir, '01 - Historical Activity reports', newName))
        else:
            move(os.path.join(downloaddir, fileEnd), os.path.join(accsdir, '00 - Historical Invoices', newName))

        copy(newPath, copyPath)

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

def find_sundays_between(start: date, end: date):
    total_days = (end - start).days + 1
    sunday = 6
    all_days = [start + timedelta(days=day) for day in range(total_days)]
    return [day for day in all_days if day.weekday() is sunday]

today = date.today()


invoiceAccs = read_excel('BNP Excel Sheet.xlsx', sheet_name='Account_Fac_Freq')

for index in invoiceAccs.index:
    fac = invoiceAccs['Facility Name'][index]
    bnpName = invoiceAccs['BNP Account Name'][index]
    wiseName = invoiceAccs['Wise Account Name'][index]
    accName = invoiceAccs['AccountName'][index]

    print(f'Getting Invoice for {accName} ({fac})...')

    if invoiceAccs['BillingFreq'][index] == 'Bimonthly':
        #Assuming is ran on 1st and 16th
        if today.day < 16:
            previousMonth = today.month - 1 if today.month != 1 else 12
            previousYear = today.year - 1 if today.month == 12 else today.year
            startBimonthly = date(previousYear, previousMonth, 16)
            endBimonthly = date(previousYear, previousMonth, monthrange(previousYear, previousMonth)[1])
        elif today.day >= 16:
            startBimonthly = date(today.year, today.month, 1)
            endBimonthly = date(today.year, today.month, 15)
        '''
        startBimonthly = today + timedelta(weeks=-2, days=-1)
        endBimonthly = startBimonthly + timedelta(weeks=+2)
        ''' 
        startPeriod = startBimonthly.strftime("%m/%d/%y")
        endPeriod = endBimonthly.strftime("%m/%d/%y")
    
        wiseStart = startBimonthly.strftime("%y-%m-%d")
        wiseEnd = endBimonthly.strftime("%y-%m-%d")

        bnpInvoice = getInvoice(bnpName, fac, startPeriod, endPeriod, accName, 'Bimonthly')
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

        if today.day < 16:
            previousMonth = today.month - 1 if today.month != 1 else 12
            previousYear = today.year - 1 if today.month == 12 else today.year
            previousBillDate = date(previousYear, previousMonth, 16)
            weeklyStartPeriods = find_sundays_between(previousBillDate, date(previousYear, previousMonth, monthrange(previousYear, previousMonth)[1]))
        elif today.day >= 16:
            previousBillDate = date(today.year, today.month, 1)
            weeklyStartPeriods = find_sundays_between(previousBillDate, date(today.year, today.month, 15))

        for count, d in enumerate(weeklyStartPeriods):
            endWeekly = d + timedelta(days=+6)
            startPeriod = d.strftime("%m/%d/%y")
            endPeriod = endWeekly.strftime("%m/%d/%y")
            wiseStart = d.strftime("%y-%m-%d")
            wiseEnd = endWeekly.strftime("%y-%m-%d")

            bnpInvoice = getInvoice(bnpName, fac, startPeriod, endPeriod, accName, 'Weekly', False, '-' + str(count))
            wiseInvoice = False

            if (bnpInvoice):
                wiseInvoice = getInvoice(wiseName, fac, wiseStart, wiseEnd, accName, 'Weekly', True, '-' + str(count))
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


    else: continue

print("Invoices Downloaded!")
print("--- %s seconds ---" % (time() - start_time))
x = input("Press Enter to finish. ")