import gspread
import datetime
import re
gp = gspread.service_account(filename='./auth.json')
spreadsheet = gp.open('TestParseMyProg')
worksheetRed = spreadsheet.worksheet("красные")
worksheetYellow = spreadsheet.worksheet("желтые")
worksheetDone = spreadsheet.worksheet("выполненные")
worksheet = spreadsheet.get_worksheet(0)
column = worksheet.col_values(4)
stringSheet = worksheet.row_values(11)

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

def prohod(dataColumn):
    j=0
    for i in dataColumn:
        j = j+1
        match = validDate(i)
        if match:
            print("Match true")
            cellCoord = 'E'+str(j)
            cell = worksheet.acell(cellCoord).value
            print(cell)
            if validDate(cell):
                print("Work done")

            else:
                print("No date")
                if isItLate(i):
                    changeOfColor(cellCoord,"red")
                    print("changed red color on "+ cellCoord)
                else:
                    changeOfColor(cellCoord, "yellow")
                    print("changed yellow color on "+ cellCoord)
        else:
            print("Match False")

def copyString(fromString, startCell):
    valX = worksheet.acell(startCell).row
    valY = worksheet.acell(startCell).col
    if worksheet.acell(startCell).value == None:
        for i in fromString:
            worksheet.update_cell(valX, valY, i)
            valY = valY + 1
    else:
        print("Value of your cell is not empty")

def copyColumn(fromColumn, startCell):
    #val = worksheet.acell(startCell).value
    valX= worksheet.acell(startCell).row
    valY=worksheet.acell(startCell).col
    if worksheet.acell(startCell).value == None:
        for i in fromColumn:
            worksheet.update_cell(valX,valY,i)
            valX=valX+1
    else:
        print("Value of your cell is not empty")


def listOfSheets():
    getList = spreadsheet.worksheets()
    sheets = []
    for i in getList:
        if findWorsheet(i):
            sheets.append(i)
    return sheets


def findWorsheet(name):
    sheet = re.search(r'\d{4}', name)
    if sheet:
        return True
    else:
        return False

def copyHeadGeneral():
    Head = []
    for i in range(7):
        Head[i] = worksheet.row_values(i+1)
    return Head
def copyHeadCalenar():
    head = []
    b = worksheet.col_values(2)
    j= 0
    for i in b:
        j = j+1
        if re.search(r'Кален\w{6}',i):
            head.append(j-2)
            head.append (i)
            head.append(j+1)
    return head
def copyHeadOper():
    head = []
    b = worksheet.col_values(2)
    j=0
    for i in b:
        j = j + 1
        if re.search(r'Опера\w{6}', i):
            head.append(j - 2)
            head.append(i)
            head.append(j + 1)
    return head
def findEnd():
    c = worksheet.col_values(3)
    length = len(c)
    print("длина таблицы: " + str(length))
    End = []
    j = 0
    print(c)
    for i in c:
         j = j + 1
         if i == '':
             test = c[j-1]+c[j]+c[j+1]
             print(test)
             if match(test):
                 End.append('B'+str(j+1))
             #print("Test string: "+test)
             if j == length - 2:
                 print(End)
                 return End


def match(string):
    if string =='':
        print('Space finded')
        return True
    else:
        return False

def updateDone(string):
    worksheetDone.update('B2:F8',string)

#prohod(column)
#copyColumn(column,"H2")
#end = findEnd()
#end= match("   ")
#print(end)

general = copyHeadOper()
print(general)