import BNPauto
from datetime import date, timedelta
import os.path
from os import mkdir
from pandas import read_excel
from calendar import monthrange
from shutil import copy

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
def getInvoice(accBNP, facility, startP, endP, accName, dateName, cycle):
    try:  
        facility, invoiceName = BNPauto.exportHandle(accBNP, facility, startP, endP, accName)

        if facility == False or invoiceName == False: 
            return False
        
        newName = accName + '-' + fac + '-' + cycle + '-Invoice'
        newPath = '\\Accounts\\00 - Historical Invoices\\' + newName + '-' + dateName + '.xlsx'
        copyPath = '\\Accounts\\02 - Current Invoices\\' + newName + '-' + dateName + '.xlsx'

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
    bnpName = invoiceAccs['AccountName'][index]
    accName = bnpName

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