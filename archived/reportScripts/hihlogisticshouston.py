import time
from pandas import read_excel
from pandas import ExcelWriter
import BNPauto

'''
---- Billing Items ----
'''
def orderProcessing(arPath):
    df = read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['OUTBOUND'])]
    count = df['Shipment'].count() - 1
    return count

def photoCopies(arPath):
    df = read_excel(io=arPath, sheet_name='Accessories')
    df = df[df['ACCOUNT ITEM'].isin(['PHOTO COPIES (BLACK & WHITE)'])]
    sum = df['QTY'].sum()
    return sum

def missedAppointment(arPath):
    df = read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['MISSED APPOINTMENT'].isin(['Y'])]
    count = df['MISSED APPOINTMENT'].count() - 1
    return count

def orderEntryCharge(arPath):
    df = read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['OUTBOUND'])]
    df = df[df['SOURCE'].isin(['MANUAL'])]
    count = df['Shipment'].count() - 1
    return count

def minimumHandlingChargeperContainer(arPath):
    df = read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['INBOUND'])]
    palletqty = df['PALLET QTY'].tolist()
    count = 0
    for qty in palletqty:
        if qty <= 21:
            count += 1
    return count

try:
    start_time = time.time()

    facility = 'Houston'
    billingAcc = 'HIH LOGISTICS, INC(HIH LOGISTICS, INC - USA (TRINA SOLAR) - HOUSTON)'
    accName = 'HIHLogisticsInc'
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

    billingItemDict = {'UFUFHDDT006' : 0, 'UFUFHDOP006' : 1, 'MISSED APPOINTMENT FEE' : 2,
                       'UFUFHDOF007OFT001CB007CS000' : 4, 'UFUFHDOF007OFT001' : 5, 'PHOTO COPIES (BLACK & WHITE)' : 3}
    itemStepsList = [orderEntryCharge, orderProcessing, missedAppointment, photoCopies, minimumHandlingChargeperContainer]

    report = read_excel(reportLoc)
    description = report['Description'].tolist()
    itemName = report['ItemName'].tolist()
    itemList = []
    for count, name in enumerate(description):
        tempList = [itemName[count], name]
        itemList.append(tempList)

    for index, item in enumerate(itemList):
        
        try:
            itemStep = itemStepsList[billingItemDict[item[0]]]
        except (KeyError):
            try:
                itemStep = itemStepsList[billingItemDict[item[1]]]
            except (KeyError):
                continue

        qty = itemStep(activityLoc)

        print(f'Item: {item}, Qty: {qty}')

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
