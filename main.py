'''
def copyHeadCalenar():
    Head = []
    b = worksheet.col_values(2)
    j = 0
    for i in b:
        j = j + 1
        search = re.search(r'Кален\w{6}',i)
        if search:
            Head = copyRange(j-1, j + 1)
            return Head
def copyHeadOper():
    Head = []
    b = worksheet.col_values(2)
    j = 0
    for i in b:
        j = j + 1
        search = re.search(r'Опера\w{6}', i)
        if search:
            Head = copyRange(j - 1, j + 2)
            return Head

def match(string):
    if string =='':
        print('Space finded')
        return True
    else:
        return False

def prepareTable():
    b = worksheet.col_values(3)
    j = 0
    for i in b:
        j=j+1
        if i == '':
            worksheet.update_cell(j,4, ' ')

'''


'''
def prohod(worksheet):
    j=0
    d=0
    r=0
    y=0
    sendDone = []
    sendRed = []
    sendYellow = []
    planData= worksheet.col_values(5)
    factData = worksheet.col_values(6)
    for i in planData:
        j = j+1
        match = validDate(i)
        if match:
            print("Match true")

            cell = factData[j]
            print(cell)

            if validDate(cell):
                d=d+1
                print("Work done")
                copyString= worksheet.row_values(j)
                sendDone.append(copyString)

            else:
                print("No date")
                if isItLate(i):
                    r=r+1
                    print("changed red color on E"+ str(j))
                    copyString = worksheet.row_values(j)
                    sendRed.append(copyString)
                else:
                    y=y+1
                    print("changed yellow color on E"+ str(j))
                    copyString = worksheet.row_values(j)
                    sendYellow.append(copyString)
        else:
            print("Match False")
    print(sendDone)
    print(len(sendDone))
    k2 = len(sendDone)
    prepareTable()
    updateDoneString(sendDone, d)
    updateRedString(sendRed, r)
    updateYellowString(sendYellow, y)
'''

import gspread
import datetime
import re
import gspread_formatting as gsf
from googleapiclient import discovery
gp = gspread.service_account(filename='./auth.json')
spreadsheet = gp.open('TestParseMyProg')
worksheetRed = spreadsheet.worksheet("красные")
worksheetYellow = spreadsheet.worksheet("желтые")
worksheetDone = spreadsheet.worksheet("выполненные")
#worksheet2 = spreadsheet.worksheet('2707')
#column = worksheet.col_values(5)
#stringSheet = worksheet.row_values(11)


def findWorsheet(name):
    sheet = re.search(r'\d{4}', name)
    if sheet:
        return True
    else:
        return False


def listOfSheets():
    getList = spreadsheet.worksheets()
    sheets = []
    print(getList)
    for i in getList:
        if findWorsheet(i):
            sheets.append(i)
    return sheets


def dateTransform(data):
    if data !='None':
        day,month,year = data.split('.')
        date = datetime.date(int(year),int(month),int(day))
        return date
    else:
        print('Please,input correct date')


def dateRazn(data1,data2):
    days = data2 - data1
    return days

def redOrYellow(data):
    razn = data.days
    if int(razn) > 14:
        print("Red color")
    else:
        print("Yellow color")


def validDate(data):
    if data == None:
        data = '0'
    matchOtmen = re.search(r'Отменен|отменено|-', data)
    if matchOtmen:
        return True
    match =re.search(r'\d\d.\d\d.\d{4}',data)
    if match:
        print("date is valid")
        return True
    else:
        print("date is not valid")
        return False


def changeOfColor(coord,color):
    if color == "red":
        worksheet.format(coord, {
            "backgroundColor": {
                "red": 1.0,
                "green": 0.0,
                "blue": 0.0
            }
        }
                         )
    elif color == "yellow":
        worksheet.format(coord, {
            "backgroundColor": {
                "red": 1.0,
                "green": 1.0,
                "blue": 0.0
            }
        }
                         )
    else:
        print("color is invalid")


def isItLate(date):
    dateNow = datetime.date.today()
    datePlan = dateTransform(date)
    razn = dateNow - datePlan
    day = razn.days
    if int(day) > 14:
        print("Red color")
        return True
    else:
        print("Yellow color")
        return False


def copyString(fromString, startCell):
    valX = worksheet.acell(startCell).row
    valY = worksheet.acell(startCell).col
    if worksheet.acell(startCell).value == None:
        for i in fromString:
            worksheet.update_cell(valX, valY, i)
            valY = valY + 1
    else:
        print("Value of your cell is not empty")

def cutHead(table):
    j = 0
    head = []
    #print('Баг в таблице:')
    for i in table:
        j=j+1
        print(i)
        if i !=[]:
            matchGen = re.search(r'Генераль\w{3}', i[0])
            if matchGen:
                print('план найден в ячейке :B'+str(j-1))
                print(i[0])
                for k in range(j-5,j+3):
                    head.append(table[k])
                return head
            matchCalendar = re.search(r'Кален\w{6}', i[0])
            if matchCalendar:
                for k in range(j-1,j+4):
                    head.append(table[k])
                return head
            matchOper = re.search(r'Опер\w{6}', i[0])
            if matchOper:
                for k in range(j-1, j + 2):
                    head.append(table[k])
                return head

def prohodFirst(table):
    doneTable =[]
    redTable = []
    yellowTable = []

    j=0
    for i in table:
        j=j+1
        if len(i)>=4:
            if validDate(i[3]):
                a = i[4:]
                if a !=[]:
                    if validDate(i[4]):
                        doneTable.append(i)
                else:
                    if isItLate(i[3]):
                        redTable.append(i)
                    else:
                        yellowTable.append(i)

    print('таблица выполнено:')
    print(doneTable)
    print('желтая таблица: ')
    print(yellowTable)
    print('таблица красные: ')
    print(redTable)
    print('Ошибка')
    print(table)
    sendTableDone = cutHead(table) + doneTable
    print('Отправляю таблицу выполненные:')
    print(sendTableDone)
    lengthTableDone = len(sendTableDone) + 2
    worksheetDone.update('B3:F' + str(lengthTableDone), sendTableDone)
    sendTableYellow = cutHead(table) + yellowTable
    print('Отправляю таблицу желтые:')
    print(sendTableYellow)
    lengthTableYellow = len(sendTableYellow) + 2
    worksheetYellow.update('B3:F' + str(lengthTableYellow), sendTableYellow)
    sendTableRed = cutHead(table) + redTable
    print('Отправляю таблицу красные:')
    print(sendTableRed)
    lengthTableRed = len(sendTableRed) + 2
    worksheetRed.update('B3:F' + str(lengthTableRed), sendTableRed)
    listOfEnds = [lengthTableDone+1,lengthTableYellow+1,lengthTableRed+1]
    return listOfEnds

def prohod(table,startList):
    doneTable = []
    redTable = []
    yellowTable = []

    j = 0
    for i in table:
        j = j + 1
        if len(i) >= 4:
            if validDate(i[3]):
                a = i[4:]
                if a != []:
                    if validDate(i[4]):
                        doneTable.append(i)
                else:
                    if isItLate(i[3]):
                        redTable.append(i)
                    else:
                        yellowTable.append(i)

    print('таблица выполнено:')
    print(doneTable)
    print('желтая таблица: ')
    print(yellowTable)
    print('таблица красные: ')
    print(redTable)
    sendTableDone = cutHead(table) + doneTable
    print('Отправляю таблицу выполненные:')
    print(sendTableDone)
    lengthTableDone = len(sendTableDone) + 2
    worksheetDone.update('B' + str(startList[0]) + ':F' + str(startList[0] + lengthTableDone), sendTableDone)
    sendTableYellow = cutHead(table) + yellowTable
    print('Отправляю таблицу желтые:')
    print(sendTableYellow)
    lengthTableYellow = len(sendTableYellow) + 2
    worksheetYellow.update('B' + str(startList[1]) + ':F' + str(startList[1] + lengthTableYellow), sendTableYellow)
    sendTableRed = cutHead(table) + redTable
    print('Отправляю таблицу красные:')
    print(sendTableRed)
    lengthTableRed = len(sendTableRed) + 2
    worksheetRed.update('B' + str(startList[2]) + ':F' + str(startList[2] + lengthTableRed), sendTableRed)
    listOfEnds = [startList[0] + lengthTableDone + 1, startList[1] + lengthTableYellow + 1,
                  startList[2] + lengthTableRed + 1]
    return listOfEnds

def getStart(worksheet):
    b = worksheet.col_values(2)
    j = 0
    start = []
    for i in b:
        j = j + 1
        search = re.search(r'Генераль\w{3}', i)
        searchCalendar = re.search(r'Кален\w{6}', i)
        searchOper = re.search(r'Опер\w{6}', i)
        if search:
            print('B' + str(j - 5))
            start.append(str(j - 5))
        if searchCalendar:
            print('B' + str(j))
            start.append(str(j))
        if searchOper:
            start.append(str(j))
            print('B' + str(j))

    print(start)
    return start

def getEnd(worksheet):
    c = worksheet.col_values(4)
    print('колонка: ')
    print(c)
    j = 0
    score = 0
    list = []

    for i in c:
         j = j + 1
         if j>10:
             if i == '':
                 if c[j - 2] == '':
                     if c[j - 3] == '':
                         list.append(str(j-3))
    list.append(str(j+1))
    print(list)
    return list

def start():
    worksheet = spreadsheet.worksheet('2747')
    startList = getStart(worksheet)
    endList = getEnd(worksheet)
    genTable = worksheet.get('B' + startList[0] + ':F' + endList[0])
    calendarTable = worksheet.get('B' + startList[1] + ':F' + endList[1])
    operTable = worksheet.get('B' + startList[2] + ':F' + endList[2])
    print(genTable)
    score = prohodFirst(genTable)
    score = prohod(calendarTable, score)
    score = prohod(operTable, score)
    worksheet2 = spreadsheet.worksheet('2707')
    startList2 = getStart(worksheet2)
    endList2 = getEnd(worksheet2)
    genTable2 = worksheet2.get('B' + startList2[0] + ':F' + endList2[0])
    calendarTable2 = worksheet2.get('B' + startList2[1] + ':F' + endList2[1])
    operTable2 = worksheet2.get('B' + startList2[2] + ':F' + endList2[2])
    print(genTable2)
    score = prohod(genTable2, score)
    score = prohod(calendarTable2, score)
    score = prohod(operTable2, score)
    worksheet3 = spreadsheet.worksheet('2707-01')
    startList3 = getStart(worksheet3)
    endList3 = getEnd(worksheet3)
    genTable3 = worksheet3.get('B' + startList3[0] + ':F' + endList3[0])
    calendarTable3 = worksheet3.get('B' + startList3[1] + ':F' + endList3[1])
    operTable3 = worksheet3.get('B' + startList3[2] + ':F' + endList3[2])
    print(genTable3)
    score = prohod(genTable3, score)
    score = prohod(calendarTable3, score)
    score = prohod(operTable3, score)

#start()
list = ['2747','2707','2707-01']

def start2(list):
    score = []
    j =0
    for i in list:
        j=j+1
        if j==1:
            worksheet = spreadsheet.worksheet(i)
            startList = getStart(worksheet)
            endList = getEnd(worksheet)
            genTable = worksheet.get('B' + startList[0] + ':F' + endList[0])
            calendarTable = worksheet.get('B' + startList[1] + ':F' + endList[1])
            operTable = worksheet.get('B' + startList[2] + ':F' + endList[2])
            print(genTable)
            score = prohodFirst(genTable)
            score = prohod(calendarTable, score)
            score = prohod(operTable, score)
        if j>1:
            worksheet = spreadsheet.worksheet(i)
            startList = getStart(worksheet)
            endList = getEnd(worksheet)
            genTable = worksheet.get('B' + startList[0] + ':F' + endList[0])
            calendarTable = worksheet.get('B' + startList[1] + ':F' + endList[1])
            operTable = worksheet.get('B' + startList[2] + ':F' + endList[2])
            print(genTable)
            score = prohod(genTable,score)
            score = prohod(calendarTable, score)
            score = prohod(operTable, score)


start2(list)
'''worksheet = spreadsheet.worksheet('2747')
fmt = gsf.cellFormat(
    textFormat=gsf.textFormat(bold=True)
    )
gsf.format_cell_range(worksheet,'B2',fmt)'''
def paintTable():
    worksheet = spreadsheet.worksheet('2747')
    test = gsf.get_effective_format(worksheet, 'B2:D4')
    gsf.format_cell_range(worksheetRed, 'B2:D4', test)
    #paintHeadGen= gsf.get_effective_format(worksheet,'B7:F7')
    #gsf.format_cell_range(worksheetRed,'B7:F7',paintHeadGen)
    paintHeadGen2 = gsf.get_effective_format(worksheet, 'B8:F10')
    gsf.format_cell_range(worksheetRed,'B8:F10',paintHeadGen2)
    paintGen =  gsf.get_effective_format(worksheet,'B11:F12')
    gsf.format_cell_range(worksheetRed, 'B11:F12',paintGen)

paintTable()