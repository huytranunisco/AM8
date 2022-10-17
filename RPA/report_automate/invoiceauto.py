import BNPauto
from datetime import date, timedelta
import os.path
from pandas import read_excel
from calendar import monthrange
from shutil import copy

def getInvoice(accBNP, facility, startP, endP, accName, dateName, cycle):
    try:  
        facility, invoiceName = BNPauto.exportHandle(accBNP, facility, startP, endP, accName)

        if facility == False or invoiceName == False: 
            return False
        
        newName = accName + '-' + fac + '-' + cycle + '-Invoice'
        newPath = 'C:\\Users\\' + os.getlogin() +'\\Desktop\\Discrepancy Reports\\Accounts\\00 - Historical Invoices\\' + newName + '-' + dateName + '.xlsx'
        copyPath = 'C:\\Users\\' + os.getlogin() + '\\Desktop\\Discrepancy Reports\\Accounts\\02 - Current Invoices\\' + newName + '-' + dateName + '.xlsx'

        os.rename(invoiceName, newPath)

        copy(newPath, copyPath)

        print('Done!')

    except Exception as e:
        if hasattr(e, 'message'):
            print(e.message)

        else:
            print('An error occured at ', e.args, e.__doc__)

def find_sundays_between(start: date, end: date):
    total_days = (end - start).days + 1
    sunday = 6
    all_days = [start + timedelta(days=day) for day in range(total_days)]
    return [day for day in all_days if day.weekday() is sunday]

today = date.today()
dateFormat = today.strftime("%m-%d-%y")

invoiceAccs = read_excel('BNP Excel Sheet.xlsx', sheet_name='Account_Fac_Freq')

for index in invoiceAccs.index:
    fac = invoiceAccs['Facility Name'][index]
    bnpName = invoiceAccs['BNP Account Name'][index]
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
        dateName = startBimonthly.strftime("%m-%d-%y")

        getInvoice(bnpName, fac, startPeriod, endPeriod, accName, dateName, 'Bimonthly')

    elif invoiceAccs['BillingFreq'][index] == 'Weekly':
        if today.day < 16:
            previousMonth = today.month - 1 if today.month != 1 else 12
            previousYear = today.year - 1 if today.month == 12 else today.year
            previousBillDate = date(previousYear, previousMonth, 16)
            weeklyStartPeriods = find_sundays_between(previousBillDate, date(previousYear, previousMonth, monthrange(previousYear, previousMonth)[1]))
        elif today.day >= 16:
            previousBillDate = date(today.year, today.month, 1)
            weeklyStartPeriods = find_sundays_between(previousBillDate, date(today.year, today.month, 15))

        for d in weeklyStartPeriods:
            endWeekly = d + timedelta(days=+6)
            startPeriod = d.strftime("%m/%d/%y")
            endPeriod = endWeekly.strftime("%m/%d/%y")
            dateName = d.strftime("%m-%d-%y")

            print(f'Start: {startPeriod}, End: {endPeriod}')

            getInvoice(bnpName, fac, startPeriod, endPeriod, accName, dateName, 'Weekly')

    else: continue

print("Invoices Downloaded!")