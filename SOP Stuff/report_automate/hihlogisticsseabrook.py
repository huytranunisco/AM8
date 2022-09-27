import time
from pandas import read_excel
from pandas import ExcelWriter
import BNPauto
from sys import exit

'''
---- Billing Items ----
'''
def handlingOffloadperPallet(arPath):
    df = read_excel(arPath, 'Order & Receipt')
    df = df[df['Shipment'].isin(['INBOUND'])]
    df = df[df['Offload Type'].isin(['FORKLIFT_WITH_PALLET'])]
    sum = df['PALLET QTY'].sum()
    return sum

def outboundHandling(arPath):
    df = read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Type'].isin(['OUTBOUND'])]

    if df['Shipment'].count() == 0:
        return 0

    count = df['Shipment'].count() - 1
    return count

def incomeInitialStorage(arPath):
    df = read_excel(arPath, 'Order & Receipt')
    df = df[df['Shipment'].isin(['INBOUND'])]
    sum = df['PALLET QTY'].sum()
    return sum

def manualEntryCharge(arPath):
    df = read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['OUTBOUND'])]
    df = df[df['SOURCE'].isin(['MANUAL'])]
    count = df['Shipment'].count() - 1
    return count

def handlingOrderProcessingRG(arPath):
    df = read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['OUTBOUND'])]
    df = df[df['TYPE'].isin(['Regular Order'])]
    count = df['TYPE'].count() - 1
    return count

try:
    start_time = time.time()

    facility = 'Seabrook'
    billingAcc = 'HIH LOGISTICS, INC(HIH LOGISTICS INC -- Fontana/Inovation/Seabrook)'
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

    billingItemDict = {'UFUFHDOF007OFT001RT001' : 0, 'UFUFSTIS007PS000' : 1,
                       'STORAGE INCOME INITIAL STORAGE per Pallet, PalletSize,none; BillingGrade:SOLAR PANEL,SOLAR PANEL,' : 1,
                       'UFUFHDLD020' : 2, 'UFUFHDDT006' : 3, 'UFUFHDOP006OT002' : 4}
    itemStepsList = [handlingOffloadperPallet, incomeInitialStorage, outboundHandling, manualEntryCharge, handlingOrderProcessingRG]

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
        print(e)

        x = input('Press Enter to Exit')
        
        exit()
    else:
        print('An error occured at ',e.args,e.__doc__)
        print('An error occured at ')

        x = input('Press Enter to Exit')

        exit()
