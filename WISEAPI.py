from curses import raw
import requests as curl

facility = 'valleyview'
url = 'https://wise.logisticsteam.com/v2/' + facility + '/report-center/activity/activity-reportv2'

headers = {'authorization':'094c7f6b-d8a5-4d62-a40c-dfe4854516b3',
           'wise-company-id':'ORG-1',
           'wise-facility-id':'F1',
           'Content-Type':'application/json'}

data_raw = {"customerId":"ORG-34557",
            "titleId":None,
            "timeFrom":"2022-06-28T00:00:00",
            "timeTo":"2022-07-05T23:59:59",
            "getApiData":True,
            "apiType":"OrderAndReceiptDetail"}


r = curl.post(url, headers=headers, data=data_raw)
r.status_code