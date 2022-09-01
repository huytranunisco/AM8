import pandas as pd

# Handling Pick per Pallet
def palletPick(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Pick Task')
    sum = df['PALLETPICKQTY'].sum()

    return df, sum

# Handling Pick per Case, WeightRangePerCase,0 - 30;
def pickCase_0_to_30(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Pick Task')
    df2 = df[(df['INDIVIDUAL ITEM WGT'] >= 0) & (df['INDIVIDUAL ITEM WGT'] <= 30)]
    sum = df2['CASEPICKQTY'].sum()

    return df2, sum

# handling order processing per order, ordertype,rg;outbound shipmethod,ltl;
def orderProcessingperOrder_RG_LTL(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Order & Receipt')
    df = df[df['Ship Method'].isin(['LTL'])]
    df = df[df['RN/DN'].str.contains('DN', case=False, na=False)]
    count = df['RN/DN'].count()

    return df, count

# handling pick per piece, weightrangepereach,over 30;
def pickPiece_over30(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Pick Task')
    df2 = df[df['INDIVIDUAL ITEM WGT'] > 30]
    sum = df2['PIECEPICKQTY'].sum()

    return df2, sum

# materials & other charges-pallet charge grade b
def palletChargeGradeB(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Materials')
    df = df[df['ITEM DESCRIPTION'].str.contains('grade b', case=False, na=False)]
    sum = df['QTY'].sum()

    return df, sum

# materials & other charges-stretch wrap
def stretchWrap(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Materials')
    df = df[df['ITEM DESCRIPTION'].str.contains('stretch wrap', case=False, na=False)]
    sum = df['QTY'].sum()

    return df, sum

# storage income initial storage per pallet, locationtype,1 high;palletsize,none;
def initialStorage(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Receive Task')
    sum = df['PALLET QTY'].sum()

    return df, sum

# customized shipping documents
def customizedShippingDocs(arPath):
    df = pd.read_excel(io=arPath, sheet_name='Accessories')
    df = df[df['DESCRIPTION'].str.contains('CUSTOMIZED SHIPPING DOCUMENTS', case=False, na=False)]
    sum = df['QTY'].sum()

    return df, sum

# accessorial cancel order per order, ispicked,yes; / cancellation charge
def accessorialCancelOrder(arPath):
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
    sum = ['QTY'].sum()

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

# receive inbound palletized / initial storage – per carton
def initialStorageNZXT(arPath):
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