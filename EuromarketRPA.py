import traceback
import time
import pandas as pd
from pandas import read_excel
from pandas import ExcelWriter
import BNPauto as BNPauto

# Handling Transload per CNTR, ContainerSize,none;
def handlingTransloadperCNTR (arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['INBOUND'])]
    count = (df['Shipment'].count()) - 1

    return df, count

# Manual Order Entry
def manualOrderEntry (arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['OUTBOUND'])]
    count = (df['Shipment'].count()) - 1

    return df, count


try:
    start_time = time.time()

    billingAcc = 'Euromarket Designs, Inc.'
    accName = 'Euromarket Designs, Inc.'
    facility = 'Valley View'
    startPeriod = input('Input Start Date (MM/DD/YY): ')
    endPeriod = input('Input End Date (MM/DD/YY): ')
    user = input('Input Username for Computer: ')
    billCycle = 'Bimonthly'

    # Input file path for activity report
    activityLoc = input('Input File Path for Wise Activity Report (Remove ""): ')
    activityLoc.replace('/', '//')

    facility, invoicePath = BNPauto.exportHandle(billingAcc, facility, startPeriod, endPeriod, user)
    facility = facility.lower()
    facility = facility.replace(' ', '')
    facility = facility.capitalize()

    reportLoc = BNPauto.invoiceToReport(user, accName, facility, billCycle, invoicePath)

    billingItemDict = {'handling transload per cntr, containersize,none;': [0, 'HANDLING TRANSLOAD PER CNTR'],
                       'manual order entry': [1, 'MANUAL ORDER ENTRY']
                       }
    itemStepsList = [handlingTransloadperCNTR, manualOrderEntry ]

    report = read_excel(reportLoc)
    billingitems = report['Description'].tolist()
    for count, name in enumerate(billingitems):
        billingitems[count] = name.lower()

    writer = ExcelWriter('itemsReport.xlsx', engine='openpyxl')

    for index, item in enumerate(billingitems):

        try:
            itemStep = itemStepsList[billingItemDict[item][0]]
        except (KeyError):
            continue

        sheetName = billingItemDict[item][1]

        print(f'Item: {item}, Sheet name: {sheetName}')

        df, qty = itemStep(activityLoc)

        report['WISE Qty'][index] = qty
        # dataCopy(df, sheetName)

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
        print('An error occured at ', e.args, e.__doc__)
        print('An error occured at ')

        x = input('Press Enter to Exit')

        exit()
        