import time
from pandas import read_excel
from pandas import ExcelWriter
import BNPauto

'''
Billing Items Here:
'''


try:
    start_time = time.time()

    billingAcc = ''
    accName = ''
    facility = ''
    startPeriod = input('Input Start Date (MM/DD/YY): ')
    endPeriod = input('Input End Date (MM/DD/YY): ')
    user = input('Input Username for Computer: ')
    billCycle = 'Bimonthly'

    #Input file path for activity report
    activityLoc = input('Input File Path for Wise Activity Report (Remove ""): ')
    activityLoc.replace('/', '//')

    facility, invoicePath = BNPauto.exportHandle(billingAcc, facility, startPeriod, endPeriod, user)
    facility = facility.lower()
    facility = facility.replace(' ', '')
    facility = facility.capitalize()

    reportLoc = BNPauto.invoiceToReport(user, accName, facility, billCycle, invoicePath)

    billingItemDict = {'Create Dictionary for Comparing Item Name and Description. '}
    itemStepsList = ['Put int the function Names']

    report = read_excel(reportLoc)
    billingitems = report['Description'].tolist()
    for count, name in enumerate(billingitems):
        billingitems[count] = name.lower()

    writer = ExcelWriter('itemsReport.xlsx', engine='openpyxl')

    for index, item in enumerate(billingitems):
        
        try:
            itemStep = itemStepsList[billingItemDict[item]]
        except (KeyError):
            continue

        print(f'Item: {item}')

        df, qty = itemStep(activityLoc)

        report['WISE Qty'][index] = qty

    report.to_excel(reportLoc, index=False)
    print('DONE!')
    print("--- %s seconds ---" % (time.time() - start_time))

    x = input('Press Enter to Exit')

    exit()

except Exception as e:
    if hasattr(e, 'message'):
        print(e.message)

        x = input('Press Enter to Exit')
        
        exit()
    else:
        print('An error occured at ',e.args,e.__doc__)
        print('An error occured at ')

        x = input('Press Enter to Exit')

        exit()
