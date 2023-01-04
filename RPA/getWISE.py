import WISEauto
import pandas as pd
from datetime import date, timedelta
import os

if __name__ == '__main__':
    downloaddir = 'C:\\Users\\' + os.getlogin() + '\\Downloads'
    accsdir = 'C:\\Users\\' + os.getlogin() +'Desktop\\Valley View - Wise Activity Reports 2022\\Jan'

    customers = pd.read_excel('customerlist.xlsx', sheet_name='Sheet1')

    customers = customers['Customer'].to_list()

    start = date(2022, 1, 1)
    end = date(2022, start.month + 1, 1) + timedelta(days=-1)
    if start.month == 12:
        end = date(2022, 12, 31)

    for customer in customers:
        WISEauto.exportReport(customer, 'Valley View', start, end)

        fileList = list(filter(os.path.isfile, glob(downloaddir + '\\*.xlsx')))
        fileList.sort(key=lambda x: os.path.getmtime(x))
        file = fileList[len(fileList) - 1]
        fileEnd = os.path.basename(file)

        move(os.path.join(downloaddir, fileEnd), os.path.join(accsdir, fileEnd))