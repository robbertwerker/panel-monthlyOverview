import xlsxwriter
from datetime import datetime
import titles
import monthConvert
import pprint
import invitesCompletes
import spssData

pp = pprint.PrettyPrinter(indent=4)
month = input('Please enter the month number, e.g. January = 1, February = 2, March = 3, etc...: \n')
while True:
    if month not in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]:
        print('Please enter a number between 1 and 12 (there are 12 months in a year)')
    else:
        break
year = input("Enter the year\n")
monthTwoDigits = '0' + month if len(month) == 1 else month
monthWritten = monthConvert.convert(monthTwoDigits)
workbook = xlsxwriter.Workbook('monthlyOverview.xlsx')
worksheet = workbook.add_worksheet('robbert')
# panel = ['Pollland: 2Orange Buddies NL: 39\nOrange Buddies BE: 49\nOrange Buddies UK: 45\nEuroQuestions: 54\nClubdoel: 31]
panel = ["Pollland","Orange Buddies NL", "Orange Buddies BE", "Orange Buddies UK", "EuroQuestions","Clubdoel"]
panelCodes = ["2", "39", "31", "54", "49","45"]


def main():
    fillStatic()
    fillInviteComplete(panelCodes)
    fillSPSSData()
    referenceCells()
    workbook.close()

def fillStatic():
    #fill titles in first column
    firstRow = titles.titlesA(monthWritten)
    row = 0
    col = 0
    for item in firstRow:
        worksheet.write(row, col, item)
        row +=1

    #fill panelnames
    panels = titles.panelNames()
    row = 1
    column = 2
    for item in panels:
        worksheet.write(row, column, item)
        column += 1

def fillInviteComplete(panelCodes):
    aggregate = []
    for panel in panelCodes:
        panelResults = invitesCompletes.main(panel, monthTwoDigits, year)
        aggregate.append(panelResults)
    pp.pprint(aggregate)
    row = 16
    col = 2
    for panel in aggregate:
        worksheet.write(row, col, panel["invited"])
        worksheet.write(row + 1, col, panel["completed"])
        col += 1

def fillSPSSData():
    spssOutput = spssData.main()
    data = spssOutput[0]
    order = spssOutput[1]
    pprint.pprint(data)

    panelNameCheck = ["Poll", "NL", "Clubdoel", "EQ", "BE", "UK"]
    orderDict = {}
    # print(order)
    for i, csv in enumerate(order):

        for panel in panelNameCheck:
            # print(panel)
            if panel in csv:
                print('panel', panel)
                orderDict[panel] = order.index(csv)
    # print(orderDict)
    orderedPanelData = []
    orderedPanelData.append(data[orderDict["Poll"]])       
    orderedPanelData.append(data[orderDict["NL"]])       
    orderedPanelData.append(data[orderDict["Clubdoel"]])       
    orderedPanelData.append(data[orderDict["EQ"]])       
    orderedPanelData.append(data[orderDict["BE"]])       
    orderedPanelData.append(data[orderDict["UK"]])       
    # print(orderedPanelData)
    row = 2
    col = 2
    for panel in orderedPanelData:
        worksheet.write(row, col,panel["IngeschrevenAfgelopenJaar"])
        worksheet.write(row+1, col,panel["IngeschrevenAfgelopenHalfJaar"])
        worksheet.write(row+2, col,panel["IngeschrevenAfgelopenKwartaal"])
        worksheet.write(row+3, col,panel["IngeschrevenAfgelopenMaand"])
        worksheet.write(row+4, col,panel["UitgeschrevenAfgelopenMaand"])
        worksheet.write(row+5, col,panel["IngeschrevenUitgeschrevenAfgelopenMaand"])
        worksheet.write(row+6, col,panel["ACTIEFDEELjaar"])
        worksheet.write(row+7, col,panel["ACTIEFDEELkwartaal"])
        worksheet.write(row+8, col,panel["ACTIEFDeelmaand"])
        col += 1
    return   

def referenceCells():
    #INACTIEFDEELtotaal
    worksheet.write('C13', '=C12-C9')
    worksheet.write('D13', '=D12-D9')
    worksheet.write('E13', '=E12-E9')
    worksheet.write('F13', '=F12-F9')
    worksheet.write('G13', '=G12-G9')
    worksheet.write('H13', '=H12-H9')

    #maand
    worksheet.write('C21', '=C17/C11')
    worksheet.write('C22', '=C18/C11')
    worksheet.write('D21', '=D17/D11')
    worksheet.write('D22', '=D18/D11')
    worksheet.write('E21', '=E17/E11')
    worksheet.write('E22', '=E18/E11')
    worksheet.write('F21', '=F17/F11')
    worksheet.write('F22', '=F18/F11')
    worksheet.write('G21', '=G17/G11')
    worksheet.write('G22', '=G18/G11')
    worksheet.write('H21', '=H17/H11')
    worksheet.write('H22', '=H18/H11')

    #jaar
    worksheet.write('C25', '=C17/C9')
    worksheet.write('C26', '=C18/C9')
    worksheet.write('D25', '=D17/D9')
    worksheet.write('D26', '=D18/D9')
    worksheet.write('E25', '=E17/E9')
    worksheet.write('E26', '=E18/E9')
    worksheet.write('F25', '=F17/F9')
    worksheet.write('F26', '=F18/F9')
    worksheet.write('G25', '=G17/G9')
    worksheet.write('G26', '=G18/G9')
    worksheet.write('H25', '=H17/H9')
    worksheet.write('H26', '=H18/H9')

    #ad hoc
    worksheet.write('B18', '=SUM(C18:H18)')
if __name__ == '__main__':
    main()