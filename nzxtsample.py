import time
import pandas as pd
import BNPauto

start_time = time.time()

billingAcc = 'NZXT'
accName = 'NZXT'
facility = 'Valley View'
startPeriod = input('Input Start Date: ')
endPeriod = input('Input End Date: ')
user = input('Input Username for Computer: ')
billCycle = 'Bimonthly'

facility, invoicePath = BNPauto.exportHandle(billingAcc, facility, startPeriod, endPeriod, user)
facility = facility.lower()
facility = facility.replace(' ', '')
facility = facility.capitalize()

reportLoc = BNPauto.invoiceToReport(user, accName, facility, billCycle, '19064728')

#Input file path for activity report
activityLoc = input('Input File Path for Wise Activity Report: ')

def dataCopy(dataframe, sheetName):
    dataframe.to_excel(writer, sheet_name = sheetName, index = False)

    writer.save()

# accessorial cancel order per order, ispicked,yes; / cancellation charge
def cancelledOrder(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['OUTBOUND'])]
    df = df[df['STATUS'].isin(['CANCELLED'])]
    count = df['STATUS'].count()

    return df, count

# outbound handling ds : over 750 cartons
def outboundHandlingDS(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['OUTBOUND'])]
    df = df[df['TYPE'].isin(['DropShip Order'])]
    sum = df['CS QTY'].sum()

    return df, sum

# routing
def routing(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Accessories')
    df = df[df['ACCOUNT ITEM'].isin(['ROUTING'])]
    sum = df['QTY'].sum()

    return df, sum

# receive inbound per carton
def inboundPerCarton(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['INBOUND'])]
    df = df[df['Offload Type'].isin(['BY_HAND_NO_PALLET'])]
    sum = df['CS QTY'].sum()

    return df, sum

# receive inbound palletized
def inboundPalletized(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['RN/DN'].str.contains('RN', case=False, na=False)]
    df = df[df['Offload Type'].isin(['FORKLIFT_WITH_PALLET'])]
    sum = df['PALLET QTY'].sum()

    return df, sum

# handling pick per pallet, ordertype,rg; / pallet pick (regular order)
def pickPerPalletRG(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Pick Task')
    df = df[df['TYPE'].isin(['Regular Order'])]
    sum = df['PALLETPICKQTY'].sum()

    return df, sum

# storage income initial storage per case / initial storage – per carton
def initialStorage(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['INBOUND'])]
    sum = df['CS QTY'].sum()

    return df, sum

# handling order processing per order, ordertype,rg;
def handlingOrderProcessingRG(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['STATUS'].isin(['SHIPPED'])]
    df = df[df['TYPE'].isin(['Regular Order'])]
    count = df['TYPE'].count()

    return df, count

# case pick (amazon order) / amazon labeling
def casePickAmazon(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['OUTBOUND'])]
    df = df[df['TYPE'].isin(['Regular Order', 'DropShip Order'])]
    df = df[df['RETAILER'].isin(['Amazon'])]
    sum = df['CS QTY'].sum()

    return df, sum

# case pick if 30lbs or less (regular order)
def casePick30lbsOrLess(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Pick Task')
    df = df[df['TYPE'].isin(['Regular Order'])]
    df = df[~df['SHIPTO'].isin(['Amazon', 'Amazon.com Services INC.'])]
    sum = df['CASEPICKQTY'].sum() + df['INNERPICKQTY'].sum() + df['PIECEPICKQTY'].sum()

    return df, sum

# rush order
def rushOrder(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Shipment'].isin(['OUTBOUND'])]
    df = df[df['IS RUSH'].isin(['Y'])]
    count = df['IS RUSH'].count()

    return df, count

# accessorial noaction per unit, materialtype,stretch wrap;
def stretchWrap(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Materials')
    df = df[df['ITEM DESCRIPTION'].str.contains('stretch wrap', case=False, na=False)]
    sum = df['QTY'].sum()

    return df, sum

# accessorial noaction per unit, materialtype,grade b pallet (40 x 48);
def gradeBPallet(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Materials')
    df = df[df['ITEM DESCRIPTION'].str.contains('Grade B Pallet', case=False, na=False)]
    sum = df['QTY'].sum()

    return df, sum

billingItemDict = {'accessorial cancel order per order, ispicked,yes;' : [0, 'CANCELLED ORDER'], 'cancellation charge' : [0, 'CANCELLED ORDER'], 
                   'outbound handling ds : over 750 cartons' : [1, 'OUTBOUND HANDLING DS'], 'routing' : [2, 'ROUTING'], 'receive inbound per carton' : [3, 'INBOUND PER CARTON'],
                   'receive inbound palletized' : [4, 'INBOUND PALLETIZED'], 'handling pick per pallet, ordertype,rg;' : [5, 'PICK PER PALLET RG'], 
                   'pallet pick (regular order)' : [5, 'PICK PER PALLET RG'], 'storage income initial storage per case' : [6, 'INITIAL STORAGE PER CASE'],
                   'initial storage - per carton' : [6, 'INITIAL STORAGE PER CARTON'], 'handling order processing per order, ordertype,rg;' : [7, 'HANDLING ORDER PROCESSING RG'],
                   'case pick (amazon order)' : [8, 'CASE PICK AMAZON'], 'amazon labeling' : [8, 'CASE PICK AMAZON'], 'case pick if 30lbs or less (regular order)' :
                   [9, 'CASE PICK 30LBS OR LESS'], 'rush order' : [10, 'RUSH ORDER'], 'accessorial noaction per unit, materialtype,stretch wrap;' : [11, 'STRETCH WRAP'],
                   'accessorial noaction per unit, materialtype,grade b pallet (40 x 48);' : [12, 'GRADE B PALLET']}
itemStepsList = [cancelledOrder, outboundHandlingDS, routing, inboundPerCarton, inboundPalletized, pickPerPalletRG, initialStorage, handlingOrderProcessingRG, 
                 casePickAmazon, casePick30lbsOrLess, rushOrder, stretchWrap, gradeBPallet]

report = pd.read_excel(reportLoc)
billingitems = report['Description'].tolist()
for count, name in enumerate(billingitems):
    billingitems[count] = name.lower()

writer = pd.ExcelWriter('itemsReport.xlsx', engine='openpyxl')

for index, item in enumerate(billingitems):
    
    try:
        itemStep = itemStepsList[billingItemDict[item][0]]
    except (KeyError):
        continue

    sheetName = billingItemDict[item][1]

    print(f'Item: {item}, Sheet name: {sheetName}')

    df, qty = itemStep(activityLoc)

    report['WISE Qty'][index] = qty
    dataCopy(df, sheetName)


report.to_excel(reportLoc, index=False)
print('DONE!')
print("--- %s seconds ---" % (time.time() - start_time))