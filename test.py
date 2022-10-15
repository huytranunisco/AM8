import BNPauto
import os.path
from os import mkdir
from pandas import read_excel
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta

'''
accountdict = {'Jetson': ['Valley View', 'JETSON(JETSON)'], 
               'International Textile & Apparel': ['Savannah', 'INTERNATIONAL TEXTILE & APPAREL, INC.(INTERNATIONAL TEXTILE & APPAREL)'], 
               'TPV': ['Valley View', 'TPV USA(TPV USA)'], 'EPI': ['Valley View', 'ENVISION PERIPHERALS INC.(EPI Non-Bonded)'],
               'Planahead': ['Fontana', 'PLANAHEAD(PLANAHEAD)'], 'Surf LLC': ['Lakewood', 'SURF 9 LLC(SURF 9 LLC)'], 
               'Intradeco Ivory': ['New Jersey', 'INTRADECO-IVORY(INTRADECO-IVORY)'], 'Intradeco Knothe': ['Morgan Lakes', 'INTRADECO-KNOTHE(INTRADECO-KNOTHE)'],
               'Delta Electronics': ['Valley View', 'DELTA ELECTRONICS (AMERICAS) LTD - NEW(DELTA ELECTRONICS (AMERICAS) LTD - NEW)'],
               'NZXT': ['Valley View', 'NZXT(NZXT)'], 'Pepsico': ['Tacoma', 'PEPSICO(PEPSICO)'],
               'Sunrise Global Marketing': ['Charleston', 'SUNRISE GLOBAL MARKETING, LLC(SUNRISE GLOBAL MARKETING, LLC)'],
               'Sunrise Global Marketing': ['Tacoma', 'SUNRISE GLOBAL MARKETING, LLC(SUNRISE GLOBAL MARKETING, LLC)'],
               'Lee Kum Kee': ['Valley View', 'LEE KUM KEE (U.S.A.) INC.(LEE KUM KEE (U.S.A.) INC.)'],'Aterian': ['Savannah', 'ATERIAN GROUP, INC.(ATERIAN GROUP, INC.)'],
               'Aterian': ['GREENWOOD', 'ATERIAN GROUP, INC.(ATERIAN GROUP, INC.)'],
               'Vita Coco': ['Valley View', 'ALL MARKET INC / VITA COCO(ALL MARKET INC / VITA COCO)'],
               'Vita Coco': ['KENT', 'ALL MARKET INC / VITA COCO(ALL MARKET INC / VITA COCO)'],
               'Vita Coco': ['Morgan Lakes', 'ALL MARKET INC / VITA COCO(ALL MARKET INC / VITA COCO)'],
               'Vita Coco': ['Via Baron', 'ALL MARKET INC / VITA COCO(ALL MARKET INC / VITA COCO)'], 'Vizio': ['Indiana', 'VIZIO(VIZIO)'],
               'Galanz': ['Innovation', 'GALANZ AMERICAS LIMITED COMPANY(GALANZ AMERICAS LIMITED COMPANY)'], 'Castlery USA': ['KENT', 'CASTLERY USA(CASTLERY USA)'],
               'Luna Wellness': ['Valley View', 'Luna Wellness LLC(Luna Wellness LLC)'],
               'HIH Logistics': ['Seabrook', 'HIH LOGISTICS, INC(HIH LOGISTICS INC -- Fontana/Inovation/Seabrook)'],
               'HIH Logistics': ['Innovation', 'HIH LOGISTICS, INC(HIH LOGISTICS INC -- Fontana/Inovation/Seabrook)'],
               'Euromarket': ['Valley View', 'Euromarket Designs, Inc.(Euromarket Designs, Inc.)'], 'July & Co PTY LTD': ['Grand Prairie', 'JULY & CO PTY LTD(JULY & CO PTY LTD)'],
               'E&S International Enterprise': ['New Jersey', 'E&S INTERNATIONAL ENTERPRISES, INC(E&S INTERNATIONAL ENTERPRISES, INC)'],
               'Banyan International': ['Indiana', 'BANYAN INTERNATIONAL(BANYAN INTERNATIONAL)'], 'Outer Inc': ['Quality-4400', 'OUTER, INC.(OUTER, INC.)'],
               'Outer Inc': ['Willow', 'OUTER, INC.(OUTER, INC.)'], 'Luna Wellness': ['New Jersey', 'Luna Wellness LLC(Luna Wellness LLC)'],
               'Sunpower Corporation': ['Innovation', 'SUNPOWER CORPORATION(SUNPOWER CORPORATION)'], 'Lecta Condat': ['New Jersey', 'LECTA - CONDAT(LECTA - CONDAT)'],
               'Turtle Beach': ['Joliet', 'TURTLE BEACH(TURTLE BEACH)'],
               'THE FIFTY/FIFTY GROUP': ['Valley View', 'THE FIFTY/FIFTY GROUP, INC DBA LOLA PRODUCTS(THE FIFTY/FIFTY GROUP, INC DBA LOLA PRODUCTS)'],
               'Deer Stags': ['KENT', 'DEER STAGS(DEER STAGS)'],
               'Safe Catch': ['KENT', 'SAFE CATCH, INC.(SAFE CATCH, INC.)'] , 'Safe Catch': ['Via Baron', 'SAFE CATCH, INC.(SAFE CATCH, INC.)'],
               'Uline': ['Lakewood', 'ULINE, INC.(ULINE)'], 'Delta Electronics': ['Murphy', 'DELTA ELECTRONICS (AMERICAS) LTD - NEW(DELTA ELECTRONICS (AMERICAS) LTD - NEW)'],
               'Radio Flyer': ['KENT', 'RADIO FLYER(RADIO FLYER)'], 'SG Companies': ['Valley View', 'THE SG COMPANIES(THE SG COMPANIES)'],
               'Circuit City Corporation': ['KENT', 'CIRCUIT CITY CORPORATION INC(CIRCUIT CITY CORPORATION INC)'],
               'ToughBuilt': ['Morgan Lakes', 'TOUGHBUILT INDUSTRIES, INC.(TOUGHBUILT INDUSTRIES, INC.)'], 'Jsonic': ['Redbluff', 'JSONIC SERVICES INC(JSONIC SERVICES INC)'],
               'Tytus Grill': ['Morgan Lakes', 'TYTUS GRILLS, LLC(TYTUS GRILLS, LLC)'], 'CoolerMaster': ['Valley View', 'COOLER MASTER(COOLER MASTER)'],
               'SCHC Wilson Art': ['New Jersey', 'SCHC - Wilsonart(SCHC - Wilsonart)'],
               'TCU Trading': ['GREENWOOD', 'TCU TRADING LTD(TCU TRADING LTD)'], 'Tytus Grill': ['Valley View', 'TYTUS GRILLS, LLC(TYTUS GRILLS, LLC)']}
'''

today = date.today()
dateFormat = today.strftime("%m-%d-%y")
folderName = "Invoices " + dateFormat
try:
    mkdir(folderName)
except (FileExistsError):
    print('Invoice folder already exists. Continuing...')

invoiceAccs = read_excel('BNP Excel Sheet.xlsx', sheet_name='Account_Fac_Freq', engine='openpyxl')

bimonthlyAccs = invoiceAccs[invoiceAccs['BillingFreq'] == 'Bimonthly']
weeklyAccs= invoiceAccs[invoiceAccs['BillingFreq'] == 'Weekly']
monthlyAccs= invoiceAccs[invoiceAccs['BillingFreq'] == 'Monthly']
dailyAccs= invoiceAccs[invoiceAccs['BillingFreq'] == 'Daily']

# intended to be ran on 1st of month and 16th of month

for index in bimonthlyAccs.index:
    fac = bimonthlyAccs['Facility Name'][index]
    bnpName = bimonthlyAccs['AccountName'][index]

    try:

        billingAcc = bnpName
        accName = bimonthlyAccs['AccountName'][index]
        facility = fac
        startBimonthly = date.today() + timedelta(weeks=-2, days=-1)
        endBimonthly = startBimonthly + timedelta(weeks=+2)
        startPeriod = startBimonthly.strftime("%m/%d/%y")
        endPeriod = endBimonthly.strftime("%m/%d/%y")

        facility, invoiceName = BNPauto.exportHandle(billingAcc, facility, startPeriod, endPeriod, accName)

        if facility == False or invoiceName == False:
            continue

        oldPath = 'Invoice[' + invoiceName + '].xlsx'
        newPath = folderName + '//' + fac + '_' +invoiceName +'.xlsx'
        os.rename(oldPath, newPath)

    except Exception as e:
        if hasattr(e, 'message'):
            print(e.message)

        else:
            print('An error occured at ', e.args, e.__doc__)

# intended to be ran on Sundays

for index in weeklyAccs.index:
    fac = weeklyAccs['Facility Name'][index]
    bnpName = weeklyAccs['AccountName'][index]

    try:

        billingAcc = bnpName
        accName = weeklyAccs['AccountName'][index]
        facility = fac
        startWeekly = date.today() + timedelta(weeks=-1, )
        endWeekly = startWeekly + timedelta(days=+6)
        startPeriod = startWeekly.strftime("%m/%d/%y")
        endPeriod = endWeekly.strftime("%m/%d/%y")

        facility, invoiceName = BNPauto.exportHandle(billingAcc, facility, startPeriod, endPeriod, accName)

        if facility == False or invoiceName == False:
            continue

        oldPath = 'Invoice[' + invoiceName + '].xlsx'
        newPath = folderName + '//' + fac + '_' +invoiceName +'.xlsx'
        os.rename(oldPath, newPath)

    except Exception as e:
        if hasattr(e, 'message'):
            print(e.message)

        else:
            print('An error occured at ', e.args, e.__doc__)

# intended to be ran on the 1st of the month

for index in monthlyAccs.index:
    fac = monthlyAccs['Facility Name'][index]
    bnpName = monthlyAccs['AccountName'][index]

    try:

        billingAcc = bnpName
        accName = monthlyAccs['AccountName'][index]
        facility = fac
        startMonthly = date.today() + relativedelta(months=-1)
        endMonthly = date.today() + timedelta(days=-1)
        startPeriod = startMonthly.strftime("%m/%d/%y")
        endPeriod = endMonthly.strftime("%m/%d/%y")

        facility, invoiceName = BNPauto.exportHandle(billingAcc, facility, startPeriod, endPeriod, accName)

        if facility == False or invoiceName == False:
            continue

        oldPath = 'Invoice[' + invoiceName + '].xlsx'
        newPath = folderName + '//' + fac + '_' +invoiceName +'.xlsx'
        os.rename(oldPath, newPath)

    except Exception as e:
        if hasattr(e, 'message'):
            print(e.message)

        else:
            print('An error occured at ', e.args, e.__doc__)

# intended to be ran daily

for index in dailyAccs.index:
    fac = dailyAccs['Facility Name'][index]
    bnpName = dailyAccs['AccountName'][index]

    try:

        billingAcc = bnpName
        accName = dailyAccs['AccountName'][index]
        facility = fac
        startDaily = date.today() + timedelta(days=-2)
        endDaily = date.today() + timedelta(days=-1)
        startPeriod = startDaily.strftime("%m/%d/%y")
        endPeriod = endDaily.strftime("%m/%d/%y")

        facility, invoiceName = BNPauto.exportHandle(billingAcc, facility, startPeriod, endPeriod, accName)

        if facility == False or invoiceName == False:
            continue

        oldPath = 'Invoice[' + invoiceName + '].xlsx'
        newPath = folderName + '//' + fac + '_' +invoiceName +'.xlsx'
        os.rename(oldPath, newPath)

    except Exception as e:
        if hasattr(e, 'message'):
            print(e.message)

        else:
            print('An error occured at ', e.args, e.__doc__)