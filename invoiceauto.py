import traceback
import time
import pandas as pd
from pandas import read_excel
from pandas import ExcelWriter
import BNPauto as BNPauto


accountdict = {'Jetson': 'Valley View', 'International Textile & Apparel': 'Savannah', 'TPV': 'Valley View', 'EPI': 'Valley View',
               'Planahead': 'Fontana', 'Surf LLC': 'Lakewood', 'Intradeco Ivory': 'New Jersey', 'Intradeco Knothe': 'Morgan Lakes',
               'Delta Electronics': 'Valley View', 'NZXT': 'Valley View', 'Pepsico': 'Tacoma',
               'Sunrise Global Marketing': 'Charleston', 'Sunrise Global Marketing': 'Tacoma', 'Lee Kum Kee': 'Valley View','Aterian': 'Savannah',
               'Aterian': 'Greenwood', 'Vita Coco': 'Valley View', 'Vita Coco': 'Kent', 'Vita Coco': 'Morgan Lakes',
               'Vita Coco': ' Via Baron', 'Vizio': 'Indiana', 'Galanz':'Innovation', 'Castlery USA':'Kent',
               'Luna Wellness':'Valley View', 'HIH Logistics': 'Seabrook', 'HIH Logistics':'Innovation',
               'Euromarket':'Valley View', 'July & Co PTY LTD':'Grand Prairie','E&S International Enterprise':'New Jersey',
               'Banyan International':'Indiana','Outer Inc':'Quality-4400', 'Outer Inc':'Willow',
               'Luna Wellness':'New Jersey','Sunpower Corporation':'Innovation','Lecta Condat':'New Jersey',
               'Turtle Beach':'Joliet', 'THE FIFTY/FIFTY GROUP':'Valley View', 'Deer Stags':'Kent',
               'Safe Catch':'Kent','Safe Catch':'Baron', 'Uline':'Lakewood', 'Delta Electronics': 'Murphy','Radio Flyer':'Kent',
               'SG company':'Valley View', 'Circuit City Corporation':'Kent','ToughBuilt':'Morgan Lakes','Jsonic':'Redbluff',
               'Tytus Grill':'Morgan Lakes', 'CoolerMaster':'Valley View','SCHC wilson art':'New Jersey',
               'TCU Trading':'Greenwood', 'Tytus Grill':'Valley View'
                                                                                            }
for key in accountdict:
    val = accountdict[key]

    try:
        start_time = time.time()

        billingAcc = key
        accName = key
        facility = val
        startPeriod = '09/16/22'
        endPeriod = '09/30/22'
        user = input('Input Username for Computer: ')
        billCycle = 'Bimonthly'

        # Input file path for activity report
        activityLoc = 'https://wise.logisticsteam.com/v2/#/rc/operation/inventory/activity-report-v2/list'
        activityLoc.replace('/', '//')

        facility, invoicePath = BNPauto.exportHandle(billingAcc, facility, startPeriod, endPeriod, user)
        facility = facility.lower()
        facility = facility.replace(' ', '')
        facility = facility.capitalize()

        reportLoc = BNPauto.invoiceToReport(user, accName, facility, billCycle, invoicePath)

        # billingItemDict = {'handling transload per cntr, containersize,none;': [0, 'HANDLING TRANSLOAD PER CNTR'],
        #                    'manual order entry': [1, 'MANUAL ORDER ENTRY']
        #                    }
        # itemStepsList = [handlingTransloadperCNTR, manualOrderEntry]
        #
        # report = read_excel(reportLoc)
        # billingitems = report['Description'].tolist()
        # for count, name in enumerate(billingitems):
        #     billingitems[count] = name.lower()
        #
        # writer = ExcelWriter('itemsReport.xlsx', engine='openpyxl')
        #
        # for index, item in enumerate(billingitems):
        #
        #     try:
        #         itemStep = itemStepsList[billingItemDict[item][0]]
        #     except (KeyError):
        #         continue
        #
        #     sheetName = billingItemDict[item][1]
        #
        #     print(f'Item: {item}, Sheet name: {sheetName}')
        #
        #     df, qty = itemStep(activityLoc)
        #
        #     report['WISE Qty'][index] = qty
        #     # dataCopy(df, sheetName)
        #
        # report.to_excel(reportLoc, index=False)
        # print('DONE!')
        # print("--- %s seconds ---" % (time.time() - start_time))

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
