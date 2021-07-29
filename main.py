import gspread
import datetime
import re

gp = gspread.service_account(filename='./auth.json')
spreadsheet = gp.open('TestParseMyProg')
worksheetRed = spreadsheet.worksheet("красные")
worksheetYellow = spreadsheet.worksheet("желтые")
worksheetDone = spreadsheet.worksheet("выполненные")
worksheet = spreadsheet.worksheet('2747')
column = worksheet.col_values(5)
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
    testTwoTable(sheets)
    return sheets

def testTwoTable(sheets):
    listOfSheets = []
    listOfSheets = sheets
    for i in listOfSheets:
        testWorksheet = spreadsheet.get_worksheet(i)
        prohod(testWorksheet)



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

def copyRange(a,b):
    result = []
    for i in range(a, b+1):
        copyStr = worksheet.row_values(i)
        print(copyStr)
        result.append(copyStr)
    return result

def copyHeadGeneral():
    Head = []
    b= worksheet.col_values(2)
    j=0
    for i in b:
        j= j+1
        search = re.search(r'Генераль\w{3}',i)
        if search:
            Head = copyRange(2,j+3)
            return Head, j+3


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
def copyTable():
    print('Список стартовых ячеек получен:')
    startList= getStart()
    print('Список конечных ячеек получен:')
    endList = getEnd()
    print('Попытка копирования таблицы Генеральный план: ')
    genTable = worksheet.get('B'+startList[1]+':F'+endList[1])
    print(genTable)
    calendarTable = worksheet.get('B'+startList[2]+':F'+endList[2])
    print(calendarTable)
    operTable = worksheet.get('B'+startList[3]+':F'+endList[2])
    print(operTable)

def prohod2():
    doneTable =[]
    redTable = []
    yellowTable = []

    j=0
    for i in table:
        j=j+1
        if len(i)>=4:
            if validDate(i[4]):
                if i[5]!='':
                    doneTable.append(i)
                else:
                    if isItLate(i[4]):
                        redTable.append(i)
                    else:
                        yellowTable.append(i)




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

def updateDoneString(string, idRaw):
    worksheetDone.update('A2:F'+str(idRaw+2),string)

def updateYellowString(string, idRaw):
    worksheetYellow.update('A2:F'+str(idRaw+2),string)

def updateRedString(string, idRaw):
    worksheetRed.update('A2:F'+str(idRaw+2),string)

def getStart():
    b = worksheet.col_values(2)
    j = 0
    start = []
    for i in b:
        j = j + 1
        search = re.search(r'Генераль\w{3}', i)
        searchCalendar = re.search(r'Кален\w{6}', i)
        searchOper = re.search(r'Опер\w{6}', i)
        if search:
            print('B' + str(j - 3))
            start.append(str(j - 3))
        if searchCalendar:
            print('B' + str(j))
            start.append(str(j))
        if searchOper:
            start.append(str(j))
            print('B' + str(j))

    print(start)
    return start

def getEnd():
    c = worksheet.col_values(3)
    j = 0
    score = 0
    list = []

    for i in c:
         j = j + 1
         if i == '':

             if c[j - 2] == '':
                 if c[j - 3] == '':
                     print(str(j), c[j-1])
                     score = score + 1
                     list.append(str(j-3))
    list.append(str(j+1))
    print(list)
    return list

def getTable(inB,inF):
    gen = worksheet.get('B'+inB+':F'+inF)
    return gen

copyTable()