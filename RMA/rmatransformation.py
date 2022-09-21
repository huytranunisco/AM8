import enum
import pandas as pd

file = "C:\\Users\\kenguyen\\Desktop\\Flag_&_Anthem_Returnly_Return_Details_created_at_2022-09-20_2022-09-20_PDT.csv"

df = pd.read_csv(file)

df = df[['RMA', 'Original Order', 'Tracking Number', 'Shopper Email', 'Shipped From Name', 'Shipped From Address 1', 'Shipped From City', 'Shipped From State',
         'Shipped From Zip', 'Shipped From Country Code', 'Barcode']]

newColDict = {'ClientID':'FLAANT0001', 'Reverse Type':'Consumer Return', 'Phone':'', 'RMA Expiration Date':'', 'Ship Method':'Small Parcel', 'Shipment Carrier':'USPS',
              'BOL':'', 'ETA':'', 'Note':'', 'UPC':'', 'Buyer ID':'', 'Return Qty':1, 'UOM':'EA', 'Serial Number':''}

newColKeys = newColDict.keys()

for index, colName in enumerate(newColKeys):
    df[colName] = newColDict[colName]

df = df.rename(columns={'Original Order' : 'Reference', 'Shipped From Name' : 'Return Party', 'Shipped From Address 1' : 'Return From Address 1',
                        'Shipped From City' : 'Return From City', 'Shipped From State' : 'Return From State', 'Shipped From Zip' : 'Return From Postal Code', 
                        'Shipped From Country Code' : 'Return From Country', 'Shopper Email' : 'Email', 'Tracking Number' : 'Pro / Tracking Number', 
                        'Barcode' : 'Item Name'})
                    
df['Reference'] = df['Reference'].str.replace('#', '')

df.to_excel('Output.xlsx', index=False)