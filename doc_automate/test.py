from docx import Document

def billPeriodMenu():
    menu = 'Billing Period for Handling:\n1. Bimonthly\n2. Weekly\n3. Daily\n'
    period = input(menu)
    while period != '1' and period != '2' and period != '3':
        period = input('Choose one of the options (1-3).')
    
    options = ['bimonthly','weekly', 'daily']

    period = options[int(period) - 1]

    return period

def replaceString(paragraph, oldText, newText):
    inline = paragraph.runs
    for i in range(len(inline)):
        if oldText in inline[i].text:
            text = inline[i].text.replace(oldText, newText)
            inline[i].text = text
            break


accName = 'NZXT'
accFacility = 'Valley View'
billerName = 'Paul Darang'
billingPeriod = billPeriodMenu()

tempFac = accFacility.replace(" ", "")

oldText = {"One" : f'{accName} ({accFacility}) SOP', "Biller Name" : billerName, "billingPeriodH" : billingPeriod, "Two" : f'{accName} ({accFacility})',
           "Three" : f'{accName}-{tempFac}-{billingPeriod.capitalize()}', "accName" : accName, "accFacility" : accFacility, "billingperiodh" : billingPeriod}
oldTextList = oldText.keys()

document = Document("C:\\Users\\kenguyen\\Documents\\SOPS\\SOP Template.docx")

if billingPeriod != 'bimonthly':
    document.paragraphs[3].text = document.paragraphs[3].text.replace(" (1-15, 16-EOM)", "")
    document.paragraphs[4].text = document.paragraphs[4].text.replace(" (1-15, 16-EOM)", "")

for paragraph in document.paragraphs:
    for text in oldTextList:
        if text in paragraph.text:
            replaceString(paragraph, text, oldText[text])

'''
textChange = f'{accName} ({accFacility}) SOP'
textStyle = document.paragraphs[0].runs
document.paragraphs[0].text = textStyle[0].text.replace("Name", textChange)

document.paragraphs[2].text = document.paragraphs[2].text.replace("Biller Name", billerName)

document.paragraphs[3].text = document.paragraphs[3].text.replace("billingPeriodH", billingPeriod)

document.paragraphs[4].text = document.paragraphs[4].text.replace("billingPeriod", billingPeriod)
textChange = f'{accName} ({accFacility})'
document.paragraphs[4].text = document.paragraphs[4].text.replace("accName", textChange)

tempFac = accFacility.replace(" ", "")
textChange = f'{accName}-{tempFac}-{billingPeriod.capitalize()}'
document.paragraphs[8].text = document.paragraphs[8].text.replace("Name", textChange)

document.paragraphs[18].text = document.paragraphs[18].text.replace("accName", accName)

document.paragraphs[19].text = document.paragraphs[19].text.replace("accFacility", accFacility)

document.paragraphs[46].text = document.paragraphs[46].text.replace("accFacility", accFacility)
'''

textChange = f'{accName}-{tempFac}-{billingPeriod.capitalize()}-SOP.docx'
document.save(textChange)
