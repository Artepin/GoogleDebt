
import gspread
import datetime
import re
import gspread_formatting as gsf
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
                for k in range(j-6,j+3):
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
    worksheetDone.update('B2:F' + str(lengthTableDone), sendTableDone)
    sendTableYellow = cutHead(table) + yellowTable
    print('Отправляю таблицу желтые:')
    print(sendTableYellow)
    lengthTableYellow = len(sendTableYellow) + 2
    worksheetYellow.update('B2:F' + str(lengthTableYellow), sendTableYellow)
    sendTableRed = cutHead(table) + redTable
    print('Отправляю таблицу красные:')
    print(sendTableRed)
    lengthTableRed = len(sendTableRed) + 2
    worksheetRed.update('B2:F' + str(lengthTableRed), sendTableRed)
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
            print('B' + str(j - 6))
            start.append(str(j - 6))
        if searchCalendar:
            print('B' + str(j-1))
            start.append(str(j-1))
        if searchOper:
            start.append(str(j-1))
            print('B' + str(j-1))

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
        worksheet = spreadsheet.worksheet(i)
        startList = getStart(worksheet)
        endList = getEnd(worksheet)
        genTable = worksheet.get('B' + startList[0] + ':F' + endList[0])
        calendarTable = worksheet.get('B' + startList[1] + ':F' + endList[1])
        operTable = worksheet.get('B' + startList[2] + ':F' + endList[2])
        print(genTable)
        if j ==1:
            score = prohodFirst(genTable)
        else:
            score = prohod(genTable, score)
        score = prohod(calendarTable, score)
        score = prohod(operTable, score)

#start2(list)

def paintTable():
    worksheet = spreadsheet.worksheet('2747')
    test = gsf.get_effective_format(worksheet, 'B2:D4')
    gsf.format_cell_range(worksheetRed, 'B2:D4', test)
    paintHeadGen= gsf.get_effective_format(worksheet,'B7:F7')
    gsf.format_cell_range(worksheetRed,'B7:F7',paintHeadGen)
    paintHeadGen2 = gsf.get_effective_format(worksheet, 'B8:F10')
    gsf.format_cell_range(worksheetRed,'B8:F10',paintHeadGen2)
    paintGen =  gsf.get_effective_format(worksheet,'B11:F12')
    gsf.format_cell_range(worksheetRed, 'B11:F12',paintGen)

def test():
    worksheet = spreadsheet.worksheet('2747')
    coordWork = getStart(worksheet)
    for i in coordWork:
        i = str(int(i)+1)
    coordRed = getStart(worksheetRed)
    for i in coordRed:
        i = str(int(i)+1)
    head = gsf.get_effective_format(worksheet, 'B'+coordWork[0]+':D'+str(int(coordWork[0])+3))
    gsf.format_cell_range(worksheetRed, 'B'+coordRed[0]+':D'+str(int(coordRed[0])+3),head)
    genHead = gsf.get_effective_format(worksheet,'B'+str(int(coordWork[0])+6)+':F'+str(int(coordWork[0])+6))
    gsf.format_cell_range(worksheetRed, 'B'+str(int(coordRed[0])+6)+':F'+str(int(coordRed[0])+6),genHead)
    genData = gsf.get_effective_format(worksheet,'B'+str(int(coordWork[0])+8))
    gsf.format_cell_range(worksheetRed, 'B'+str(int(coordRed[0])+8)+':F'+str(int(coordRed[1])-1),genData)
    test = gsf.get_effective_format(worksheet,'B7:F7')
    print(test)
    '''
    test2 = gsf.get_effective_format(worksheet, 'B8:F12')
    gsf.format_cell_range(worksheetRed,'B8:F'+coordRed(1),test2)
    '''

#test()
def test2():
    worksheet = spreadsheet.worksheet('2747')
    coordWork = getStart(worksheet)
    for i in coordWork:
        i = str(int(i) + 1)
    coordRed = getStart(worksheetRed)
    for i in coordRed:
        i = str(int(i) + 1)
    j = 0
    test2 = gsf.get_effective_format(worksheet, 'B2:D4')
    gen = gsf.get_effective_format(worksheet, 'B' + coordWork[0] + ':F' + coordWork[0])
    print(gen)
    for i in coordRed:
        j=j+1
        if (j+2)%3==0:
            gsf.format_cell_range(worksheetRed, 'B'+coordRed[j]+':F'+coordRed[j],gen)
        #if (j+1)%3==0:

def test3(worksheet):
    b = worksheet.col_values(2)
    genCoord = []
    calenCoord = []
    operCoord = []
    j = 0
    for i in b:
        j = j + 1
        searchGen = re.search(r'Генераль\w{3}', i)
        searchCalendar = re.search(r'Кален\w{6}', i)
        searchOper = re.search(r'Опер\w{6}', i)
        if searchGen:
            print('B'+str(j))
            print(i)
            genCoord.append(str(j))
        if searchCalendar:
            print('B' + str(j))
            print(i)
            calenCoord.append(str(j))
        if searchOper:
            print('B' + str(j))
            print(i)
            operCoord.append(str(j))
    return genCoord,calenCoord,operCoord

def test4():
    worksheet = spreadsheet.worksheet('2747')
    genZakazCoord, calendarZakazCoord, operZakazCoord = test3(worksheet)
    genRedCoord, calendarRedCoord, operRedCoord = test3(worksheetRed)
    genYellowCoord, calendarYellowCoord, operYellowCoord = test3(worksheetYellow)
    genDoneCoord, calendarDoneCoord, operDoneCoord = test3(worksheetDone)
    j=-1
    for i in genRedCoord:
        j=j+1
        zakaz = gsf.get_effective_format(worksheet, 'C'+str(int(genZakazCoord[0])-4))
        gsf.format_cell_range(worksheetRed, 'B'+str(int(genRedCoord[j])-5)+':D'+str(int(genRedCoord[j])-3),zakaz)
        print('Координаты № заказа: B'+str(int(genRedCoord[j])-5)+':D'+str(int(genRedCoord[j])-3))
        genHeadPaint = gsf.get_effective_format(worksheet, 'B'+genZakazCoord[0]+':F'+genZakazCoord[0])
        print('Координаты Шапки ген заказа: B' + genZakazCoord[0] + ':F'+genZakazCoord[0])
        gsf.format_cell_range(worksheetRed, 'B'+genRedCoord[j]+':F'+genRedCoord[j], genHeadPaint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B'+str(int(genZakazCoord[0])+1)+':F'+str(int(calendarZakazCoord[0])-2))
        gsf.format_cell_range(worksheetRed, 'B'+str(int(genRedCoord[j])+1)+':F'+str(int(calendarRedCoord[j])-2),genDataPaint)
    k =-1
    for i in calendarRedCoord:
        k=k+1
        genHeadPaint = gsf.get_effective_format(worksheet, 'B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        gsf.format_cell_range(worksheetRed, 'B' + calendarRedCoord[k] + ':F' + calendarRedCoord[k], genHeadPaint)
        calendarHead2Paint = gsf.get_effective_format(worksheet,'B'+str(int(calendarZakazCoord[0])+2)+':F'+str(int(calendarZakazCoord[0])+4))
        print('Координаты шапки Календарного плана:'+'B'+str(int(calendarZakazCoord[0])+2)+':F'+str(int(calendarZakazCoord[0])+4))
        gsf.format_cell_range(worksheetRed, 'B'+str(int(calendarRedCoord[k])+2)+':F'+str(int(calendarRedCoord[k])+4),calendarHead2Paint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 5) + ':F' + str(int(operZakazCoord[0]) - 2))
        gsf.format_cell_range(worksheetRed, 'B' + str(int(calendarRedCoord[k]) + 5) + ':F' + str(int(operRedCoord[k]) - 2),genDataPaint)
    n = -1
    for i in operRedCoord:
        n=n+1
        genHeadPaint = gsf.get_effective_format(worksheet, 'B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        gsf.format_cell_range(worksheetRed, 'B' + operRedCoord[n] + ':F' + operRedCoord[n], genHeadPaint)
        calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        print('Координаты шапки Календарного плана:' + 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        gsf.format_cell_range(worksheetRed,'B' + str(int(operRedCoord[n]) + 1) + ':F' + str(int(operRedCoord[n]) + 2),calendarHead2Paint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 3) + ':F' + str(int(operZakazCoord[0]) + 3))
        gsf.format_cell_range(worksheetRed,'B' + str(int(operRedCoord[n]) + 3) + ':F' + str(int(operRedCoord[n]) + 3),genDataPaint)

    a = -1
    for i in genYellowCoord:
        a=a+1
        zakaz = gsf.get_effective_format(worksheet, 'C'+str(int(genZakazCoord[0])-4))
        gsf.format_cell_range(worksheetYellow, 'B'+str(int(genYellowCoord[a])-5)+':D'+str(int(genYellowCoord[a])-3),zakaz)
        print('Желтая таблица Координаты № заказа: B'+str(int(genYellowCoord[a])-5)+':D'+str(int(genYellowCoord[a])-3))
        genHeadPaint = gsf.get_effective_format(worksheet, 'B'+genZakazCoord[0]+':F'+genZakazCoord[0])
        print('Желтая таблица  Координаты Шапки ген заказа: B' + genZakazCoord[0] + ':F'+genZakazCoord[0])
        gsf.format_cell_range(worksheetYellow, 'B'+genYellowCoord[a]+':F'+genYellowCoord[a], genHeadPaint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B'+str(int(genZakazCoord[0])+1)+':F'+str(int(calendarZakazCoord[0])-2))
        gsf.format_cell_range(worksheetYellow, 'B'+str(int(genYellowCoord[a])+1)+':F'+str(int(calendarYellowCoord[a])-2),genDataPaint)

    b = -1
    for i in calendarYellowCoord:
        b = b+1
        genHeadPaint = gsf.get_effective_format(worksheet, 'B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        gsf.format_cell_range(worksheetYellow, 'B' + calendarYellowCoord[b] + ':F' + calendarYellowCoord[b], genHeadPaint)
        calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        print('Координаты шапки Календарного плана:' + 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        gsf.format_cell_range(worksheetYellow, 'B' + str(int(calendarYellowCoord[b]) + 2) + ':F' + str(int(calendarYellowCoord[b]) + 4),calendarHead2Paint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 5) + ':F' + str(int(operZakazCoord[0]) - 2))
        gsf.format_cell_range(worksheetYellow,'B' + str(int(calendarYellowCoord[b]) + 5) + ':F' + str(int(operYellowCoord[b]) - 2),genDataPaint)
    c = -1
    for i in operYellowCoord:
        c=c+1
        genHeadPaint = gsf.get_effective_format(worksheet, 'B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        gsf.format_cell_range(worksheetYellow, 'B' + operYellowCoord[c] + ':F' + operYellowCoord[c], genHeadPaint)
        calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        print('Координаты шапки Календарного плана:' + 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        gsf.format_cell_range(worksheetYellow, 'B' + str(int(operYellowCoord[c]) + 1) + ':F' + str(int(operYellowCoord[c]) + 2),calendarHead2Paint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 3) + ':F' + str(int(operZakazCoord[0]) + 3))
        gsf.format_cell_range(worksheetYellow, 'B' + str(int(operYellowCoord[c]) + 3) + ':F' + str(int(operYellowCoord[c]) + 3),genDataPaint)

    d = -1
    for i in genDoneCoord:
        d =d+1
        zakaz = gsf.get_effective_format(worksheet, 'C' + str(int(genZakazCoord[0]) - 4))
        gsf.format_cell_range(worksheetDone,'B' + str(int(genDoneCoord[d]) - 5) + ':D' + str(int(genDoneCoord[d]) - 3), zakaz)
        print('Желтая таблица Координаты № заказа: B' + str(int(genDoneCoord[d]) - 5) + ':D' + str(int(genDoneCoord[d]) - 3))
        genHeadPaint = gsf.get_effective_format(worksheet, 'B' + genZakazCoord[0] + ':F' + genZakazCoord[0])
        print('Желтая таблица  Координаты Шапки ген заказа: B' + genZakazCoord[0] + ':F' + genZakazCoord[0])
        gsf.format_cell_range(worksheetDone, 'B' + genDoneCoord[d] + ':F' + genDoneCoord[d], genHeadPaint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(genZakazCoord[0]) + 1) + ':F' + str(int(calendarZakazCoord[0]) - 2))
        gsf.format_cell_range(worksheetDone,'B' + str(int(genDoneCoord[d]) + 1) + ':F' + str(int(calendarDoneCoord[d]) - 2),genDataPaint)

    e = -1
    for i in calendarDoneCoord:
        e = e+1
        genHeadPaint = gsf.get_effective_format(worksheet, 'B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        print('Таблица выполненные координаты шапки календарного плана: B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        gsf.format_cell_range(worksheetDone, 'B' + calendarDoneCoord[e] + ':F' + calendarDoneCoord[e],genHeadPaint)
        calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        print('Таблица выполненные координаты шапки календарного плана:' + 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        gsf.format_cell_range(worksheetDone,'B' + str(int(calendarDoneCoord[e]) + 2) + ':F' + str(int(calendarDoneCoord[e]) + 4),calendarHead2Paint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 5) + ':F' + str(int(operZakazCoord[0]) - 2))
        gsf.format_cell_range(worksheetDone,'B' + str(int(calendarDoneCoord[e]) + 5) + ':F' + str(int(operDoneCoord[e]) - 2),genDataPaint)

worksheet = spreadsheet.worksheet('2747')
test4()