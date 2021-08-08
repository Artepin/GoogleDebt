
import gspread

from datetime import datetime
from datetime import date
import re
import gspread_formatting as gsf
import time
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
        if data!= '-':
            print(data)
            if len(data)<10:
                testDate = datetime.strptime(data, '%d.%m.%y')
            else:
                testDate = datetime.strptime(data, '%d.%m.%Y')
            print(data)
            #date = datetime.date(int(year),int(month),int(day))
            print(testDate)
            return testDate
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
        #print("date is valid")
        return True
    else:
        match2 = re.search(r'\d\d.\d\d.\d{2}', data)
        if match2:
            return True
        #print("date is not valid")
        return False

def isItLate(data):
    dateNow = date.today()
    datePlan = dateTransform(data)
    if datePlan ==None:
        #print(data)
        datePlan = datetime.strptime(data, '%d.%m.%y')
    razn = datePlan.date() - dateNow
    day = razn.days
    if int(day) <= 0:
        return True
    else:
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
    b = worksheet.col_values(2)
    j = 0
    start = []
    for i in b:
        j = j+1
        #search = re.search(r'Генераль\w{3}', i)
        searchCalendar = re.search(r'Кален\w{6}', i)
        searchOper = re.search(r'Опер\w{6}', i)
        searchPerep = re.search(r'Переписка', i)

        if searchCalendar:
            print('B' + str(j - 1))
            start.append(str(j - 1))
        if searchOper:
            start.append(str(j - 1))
            print('B' + str(j - 1))
        if searchPerep:
            start.append(str(j - 1))
            print('B' + str(j - 1))
    return start

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

def paintRed(list):
    for i in list:
        worksheet = spreadsheet.worksheet(i)
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

def testOPti(spreadsheet):
    gsfb = gsf.batch_updater(spreadsheet)
    worksheet = spreadsheet.worksheet('шаблоны')
    genZakazCoord, calendarZakazCoord, operZakazCoord = test3(worksheet)
    genRedCoord, calendarRedCoord, operRedCoord = test3(worksheetRed)
    genYellowCoord, calendarYellowCoord, operYellowCoord = test3(worksheetYellow)
    genDoneCoord, calendarDoneCoord, operDoneCoord = test3(worksheetDone)
    j = -1
    for i in genRedCoord:
        j = j + 1
        zakaz = gsf.get_effective_format(worksheet, 'C' + str(int(genZakazCoord[0]) - 4))
        gsfb.format_cell_range(worksheetRed, 'B' + str(int(genRedCoord[j]) - 5) + ':D' + str(int(genRedCoord[j]) - 3),zakaz)
        print('Крашу шапку генерального плана (заказ) красной таблицы: B' + str(int(genRedCoord[j]) - 5) + ':D' + str(int(genRedCoord[j]) - 3))
        genHeadPaint = gsf.get_effective_format(worksheet, 'B' + genZakazCoord[0] + ':F' + genZakazCoord[0])
        gsfb.format_cell_range(worksheetRed, 'B' + genRedCoord[j] + ':F' + genRedCoord[j], genHeadPaint)
        print('Крашу шапку генерального плана 1 красной таблицы: B' + genRedCoord[j] + ':F' + genRedCoord[j])
        genHeadPaint2 = gsf.get_effective_format(worksheet, 'B' + str(int(genZakazCoord[0])+1) + ':F' + str(int(genZakazCoord[0])+1))
        gsfb.format_cell_range(worksheetRed, 'B' + str(int(genRedCoord[j])+1)  + ':F' + str(int(genRedCoord[j])+3), genHeadPaint2)
        print('Крашу шапку генерального плана 2 красной таблицы: B' + str(int(genRedCoord[j])+1) + ':F' + str(int(genRedCoord[j])+3))
        genDataPaint = gsf.get_effective_format(worksheet, 'B'+str(int(genZakazCoord[0])+4)+':F'+str(int(genZakazCoord[0])+4))
        print(str(int(genRedCoord[j])+4)+str(int(calendarRedCoord[j])-2))
        gsfb.format_cell_range(worksheetRed,'B' + str(int(genRedCoord[j])+4) + ':F' + str(int(calendarRedCoord[j]) - 2),genDataPaint)
    time.sleep(20)

    k=-1
    for i in calendarRedCoord:
        k=k+1
        genHeadPaint = gsf.get_effective_format(worksheet, 'B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        gsfb.format_cell_range(worksheetRed, 'B' + calendarRedCoord[k] + ':F' + calendarRedCoord[k], genHeadPaint)
        calendarHead2Paint = gsf.get_effective_format(worksheet,'B'+str(int(calendarZakazCoord[0])+2)+':F'+str(int(calendarZakazCoord[0])+4))
        print('Координаты шапки Календарного плана:'+'B'+str(int(calendarZakazCoord[0])+2)+':F'+str(int(calendarZakazCoord[0])+4))
        gsfb.format_cell_range(worksheetRed, 'B'+str(int(calendarRedCoord[k])+1)+':F'+str(int(calendarRedCoord[k])+3),calendarHead2Paint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 4) + ':F' + str(int(operZakazCoord[0]) - 3))
        gsfb.format_cell_range(worksheetRed, 'B' + str(int(calendarRedCoord[k]) + 4) + ':F' + str(int(operRedCoord[k]) - 2),genDataPaint)
    time.sleep(20)
    n = -1
    for i in operRedCoord:
        n=n+1
        genHeadPaint = gsf.get_effective_format(worksheet, 'B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        gsfb.format_cell_range(worksheetRed, 'B' + operRedCoord[n] + ':F' + operRedCoord[n], genHeadPaint)
        calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        print('Координаты шапки Календарного плана:' + 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        gsfb.format_cell_range(worksheetRed,'B' + str(int(operRedCoord[n]) + 1) + ':F' + str(int(operRedCoord[n]) + 2),calendarHead2Paint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 3) + ':F' + str(int(operZakazCoord[0]) + 3))
        gsfb.format_cell_range(worksheetRed,'B' + str(int(operRedCoord[n]) + 3) + ':F' + str(int(operRedCoord[n]) + 3),genDataPaint)

    gsfb.execute()
    '''
    '''

def getTable(table):
    worksheet = spreadsheet.worksheet(table)
    b = worksheet.col_values(2)
    j = 0
    start = []
    for i in b:
        j=j+1
        searchPerep = re.search(r'Переписка', i)
        if searchPerep:
            table =worksheet.batch_get(['B2:F'+str(j)])
            #print(j)
            table2 =table[0]
            #print(table2)
    return table2

def findWords(table):
    j=0
    words = []
    for i in table:
        j = j + 1
        if i !=[]:
            searchGen = re.search(r'Генераль\w{3}', i[0])
            searchCalendar = re.search(r'Кален\w{6}', i[0])
            searchOper = re.search(r'Опер\w{6}', i[0])
            searchPerep = re.search(r'Переписка', i[0])
            if searchGen:
                #print('B' + str(j))
                #print(i)
                words.append(j)
            if searchCalendar:
                #print('B' + str(j))
                #print(i)
                words.append(j)
            if searchOper:
                #print('B' + str(j))
                #print(i)
                words.append(j)
            if searchPerep:
                #print('B' + str(j))
                #print(i)
                words.append((j))

    return words

def cutGen(table):
    words = findWords(table)
    general = words[0]
    calendar = words[1]
    genTable = table[general:calendar-1]
    #print(genTable)
    return genTable

def cutCalendar(table):
    words = findWords(table)
    calendar = words[1]
    oper = words[2]
    #print(calendar,oper)
    calendarTable = table[calendar-1:oper - 1]
    #print(calendarTable)
    return calendarTable

def cutOPer(table):
    words = findWords(table)
    oper = words[2]
    perep = words[3]
    operTable = table[oper-1:perep - 1]
    return operTable

def cutHead(table):
    j = 1
    head = []
    for i in table:
        j=j+1
        #print(i)
        if i !=[]:
            matchGen = re.search(r'Генераль\w{3}', i[0])
            if matchGen:
                #print('план найден в ячейке :B'+str(j-1))
                #print(i[0])
                for k in range(j-7,j+2):
                    head.append(table[k])
                return head
            matchCalendar = re.search(r'Кален\w{6}', i[0])
            if matchCalendar:
                for k in range(j-2,j+2):
                    head.append(table[k])
                return head
            matchOper = re.search(r'Опер\w{6}', i[0])
            if matchOper:
                for k in range(j-2, j + 2):
                    head.append(table[k])
                return head


def parse2(list):
    #isItLate('20.08.2021')
    j=-1
    #list2 = []
    #if len(list)>5:
    #    list2 = list[5:]
    redTable = []
    yellowTable = []
    doneTable = []
    for i in list:
        print(str(j+1)+'Таблица')
        j = j + 1
        table = getTable(i)
        gen = cutGen(table)
        headGenZakaz = cutHead(table)
        #headGen = cutGenHead(headGenZakaz)
        calendar = cutCalendar(table)
        headCalendar = cutHead(calendar)
        oper = cutOPer(table)
        headOper = cutHead(oper)
        doneTable1 = []
        redTable1 = []
        yellowTable1 = []
        doneTable2 = []
        redTable2 = []
        yellowTable2 = []
        doneTable3 = []
        redTable3 = []
        yellowTable3 = []
        scoreRed = 0
        scoreYellow = 0
        scoreDone = 0


        for i in gen:
            if len(i) >= 4:
                b = i[3:]
                if b!=[]:
                    if validDate(i[3]):
                        a = i[4:]
                        if a != []:
                            if validDate(i[4]):
                                doneTable1.append(i)
                        else:
                            if i[3] !='-':
                                if isItLate(i[3]):
                                    print()
                                    redTable1.append(i)
                                else:
                                    yellowTable1.append(i)
                            else:
                                doneTable1.append(i)

        a = []

        doneTable1.append(a)
        doneTable1.append(a)
        redTable1.append(a)
        redTable1.append(a)
        yellowTable1.append(a)
        yellowTable1.append(a)


        redTableGen = headGenZakaz + redTable1 + []
        yellowTableGen = headGenZakaz + yellowTable1 + []
        doneTableGen = headGenZakaz + doneTable1 + []

        for i in calendar:
            if len(i) >= 4:
                b = i[3:]
                if b != []:
                    if validDate(i[3]):
                        a = i[4:]
                        if a != []:
                            if validDate(i[4]):
                                doneTable2.append(i)
                        else:
                            if i[3] != '-':
                                if isItLate(i[3]):
                                    print()
                                    redTable2.append(i)
                                else:
                                    yellowTable2.append(i)
                            else:
                                doneTable2.append(i)

        a = []

        doneTable2.append(a)
        doneTable2.append(a)
        redTable2.append(a)
        redTable2.append(a)
        yellowTable2.append(a)
        yellowTable2.append(a)

        redTableCalendar = headCalendar + redTable2 + []
        yellowTableCalendar = headCalendar + yellowTable2 + []
        doneTableCalendar = headCalendar + doneTable2 + []

        for i in oper:
            if len(i) >= 4:
                b = i[3:]
                if b != []:
                    if validDate(i[3]):
                        a = i[4:]
                        if a != []:
                            if validDate(i[4]):
                                doneTable3.append(i)
                        else:
                            if i[3] != '-':
                                if isItLate(i[3]):
                                    print()
                                    redTable3.append(i)
                                else:
                                    yellowTable3.append(i)
                            else:
                                doneTable3.append(i)

        redTableOper = headOper+ redTable3 + []
        yellowTableOper = headOper+yellowTable3
        doneTableOper = headOper+doneTable3

        redTable += [] + redTableGen + redTableCalendar + redTableOper
        yellowTable += yellowTableGen + yellowTableCalendar + yellowTableOper
        doneTable += doneTableGen + doneTableCalendar + doneTableOper

    '''
    if list2!=[]:

        for i in list2:
            j = j + 1
            table = getTable(i)
            gen = cutGen(table)
            headGenZakaz = cutHead(table)
            headGen = cutGenHead(headGenZakaz)
            calendar = cutCalendar(table)
            headCalendar = cutHead(calendar)
            oper = cutOPer(table)
            headOper = cutHead(oper)
            doneTable1 = []
            redTable1 = []
            yellowTable1 = []
            doneTable2 = []
            redTable2 = []
            yellowTable2 = []
            doneTable3 = []
            redTable3 = []
            yellowTable3 = []
            scoreRed = 0
            scoreYellow = 0
            scoreDone = 0

            for i in gen:
                if len(i) >= 4:
                    b = i[3:]
                    if b != []:
                        if validDate(i[3]):
                            a = i[4:]
                            if a != []:
                                if validDate(i[4]):
                                    doneTable1.append(i)
                            else:
                                if i[3] != '-':
                                    if isItLate(i[3]):
                                        print()
                                        redTable1.append(i)
                                    else:
                                        yellowTable1.append(i)
                                else:
                                    doneTable1.append(i)

            a = []

            doneTable1.append(a)
            doneTable1.append(a)
            redTable1.append(a)
            redTable1.append(a)
            yellowTable1.append(a)
            yellowTable1.append(a)

            redTableGen = headGenZakaz + redTable1 + []
            yellowTableGen = headGenZakaz + yellowTable1 + []
            doneTableGen = headGenZakaz + doneTable1 + []

            for i in calendar:
                if len(i) >= 4:
                    if validDate(i[3]):
                        a = i[4:]
                        if a != []:
                            if validDate(i[4]):
                                doneTable2.append(i)
                        else:
                            if isItLate(i[3]):
                                redTable2.append(i)
                            else:
                                yellowTable2.append(i)

            a = []

            doneTable2.append(a)
            doneTable2.append(a)
            redTable2.append(a)
            redTable2.append(a)
            yellowTable2.append(a)
            yellowTable2.append(a)

            redTableCalendar = headCalendar + redTable2 + []
            yellowTableCalendar = headCalendar + yellowTable2 + []
            doneTableCalendar = headCalendar + doneTable2 + []

            for i in oper:
                if len(i) >= 4:
                    if validDate(i[3]):
                        a = i[4:]
                        if a != []:
                            if validDate(i[4]):
                                doneTable3.append(i)
                        else:
                            if isItLate(i[3]):
                                redTable3.append(i)
                            else:
                                yellowTable3.append(i)

            redTableOper = headOper + redTable3 + []
            yellowTableOper = headOper + yellowTable3
            doneTableOper = headOper + doneTable3

            redTable += [] + redTableGen + redTableCalendar + redTableOper
            yellowTable += yellowTableGen + yellowTableCalendar + yellowTableOper
            doneTable += doneTableGen + doneTableCalendar + doneTableOper
'''

    print('отправляю таблицу:')
    for i in redTable:
        print(i)
    print(len(redTable))
    #time.sleep(30)
    worksheetRed.clear()
    worksheetRed.update('B2:F' + str(len(redTable) + 2), redTable)
    #scoreRed += len(redTable) + 6
    worksheetYellow.clear()
    worksheetYellow.update('B2:F'+str(len(yellowTable) + 2),yellowTable)
    #scoreYellow+=len(yellowTable) + 3
    worksheetDone.clear()
    worksheetDone.update('B2:F'+str(len(doneTable) + 2),doneTable)
    #scoreDone+=len(doneTable)+3


    '''
    else:
        worksheetRed.update('B' + str(scoreRed+2) + ':F' + str(len(redTable) + scoreRed + 1))
        scoreRed += len(redTable) + 3
        print('счётчик красные: ' + str(scoreRed))
        worksheetYellow.update('B' + str(scoreYellow) + ':F' + str(len(yellowTable) + scoreYellow + 1))
        scoreYellow += len(yellowTable) + 3
        print('счётчик желтые: ' +str(scoreYellow))
        worksheetDone.update('B' + str(scoreDone) + ':F' + str(len(doneTable) + scoreDone + 1))
        scoreRed += len(doneTable) + 3
        print('счётчик выполненные: '+str(scoreDone))
    '''

def testOPti2(spreadsheet):
    gsfb = gsf.batch_updater(spreadsheet)
    worksheet = spreadsheet.worksheet('шаблоны')
    genZakazCoord, calendarZakazCoord, operZakazCoord = test3(worksheet)
    genRedCoord, calendarRedCoord, operRedCoord = test3(worksheetRed)
    genYellowCoord, calendarYellowCoord, operYellowCoord = test3(worksheetYellow)
    genDoneCoord, calendarDoneCoord, operDoneCoord = test3(worksheetDone)


    zakaz = gsf.get_effective_format(worksheet, 'C' + str(int(genZakazCoord[0]) - 4))
    genHeadPaint = gsf.get_effective_format(worksheet, 'B' + genZakazCoord[0] + ':F' + genZakazCoord[0])
    genHeadPaint2 = gsf.get_effective_format(worksheet, 'B' + str(int(genZakazCoord[0]) + 1) + ':F' + str(int(genZakazCoord[0]) + 1))
    genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(genZakazCoord[0]) + 4) + ':F' + str(int(genZakazCoord[0]) + 4))

    calendarHeadPaint = gsf.get_effective_format(worksheet, 'B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
    calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
    calendarDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 4) + ':F' + str(int(operZakazCoord[0]) - 3))

    operHeadPaint = gsf.get_effective_format(worksheet, 'B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
    operHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
    operDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 3) + ':F' + str(int(operZakazCoord[0]) + 3))
    j = -1
    for i in genRedCoord:
        j = j + 1
        gsfb.format_cell_range(worksheetRed, 'B' + str(int(genRedCoord[j]) - 5) + ':D' + str(int(genRedCoord[j]) - 3),zakaz)
        print('Крашу шапку генерального плана (заказ) красной таблицы: B' + str(int(genRedCoord[j]) - 5) + ':D' + str(int(genRedCoord[j]) - 3))
        gsfb.format_cell_range(worksheetRed, 'B' + genRedCoord[j] + ':F' + genRedCoord[j], genHeadPaint)
        print('Крашу шапку генерального плана 1 красной таблицы: B' + genRedCoord[j] + ':F' + genRedCoord[j])
        gsfb.format_cell_range(worksheetRed, 'B' + str(int(genRedCoord[j])+1)  + ':F' + str(int(genRedCoord[j])+3), genHeadPaint2)
        print('Крашу шапку генерального плана 2 красной таблицы: B' + str(int(genRedCoord[j])+1) + ':F' + str(int(genRedCoord[j])+3))
        print(str(int(genRedCoord[j])+4)+str(int(calendarRedCoord[j])-2))
        gsfb.format_cell_range(worksheetRed,'B' + str(int(genRedCoord[j])+4) + ':F' + str(int(calendarRedCoord[j]) - 2),genDataPaint)
    #time.sleep(20)

    k=-1
    for i in calendarRedCoord:
        k=k+1
        print('Координаты Шапки календарного плана: B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        gsfb.format_cell_range(worksheetRed, 'B' + calendarRedCoord[k] + ':F' + calendarRedCoord[k], calendarHeadPaint)
        print('Координаты шапки Календарного плана:'+'B'+str(int(calendarZakazCoord[0])+2)+':F'+str(int(calendarZakazCoord[0])+4))
        gsfb.format_cell_range(worksheetRed, 'B'+str(int(calendarRedCoord[k])+1)+':F'+str(int(calendarRedCoord[k])+3),calendarHead2Paint)
        gsfb.format_cell_range(worksheetRed, 'B' + str(int(calendarRedCoord[k]) + 4) + ':F' + str(int(operRedCoord[k]) - 2),calendarDataPaint)

    n = -1
    for i in operRedCoord:
        n=n+1
        print('Координаты Шапки календарного плана: B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        gsfb.format_cell_range(worksheetRed, 'B' + operRedCoord[n] + ':F' + operRedCoord[n], operHeadPaint)
        print('Координаты шапки Календарного плана:' + 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        gsfb.format_cell_range(worksheetRed,'B' + str(int(operRedCoord[n]) + 1) + ':F' + str(int(operRedCoord[n]) + 2),operHead2Paint)
        gsfb.format_cell_range(worksheetRed,'B' + str(int(operRedCoord[n]) + 3) + ':F' + str(int(operRedCoord[n]) + 3),operDataPaint)


    a = -1
    for i in genYellowCoord:
        a = a + 1
        #zakaz = gsf.get_effective_format(worksheet, 'C' + str(int(genZakazCoord[0]) - 4))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(genYellowCoord[a]) - 5) + ':D' + str(int(genYellowCoord[a]) - 3), zakaz)
        print('Желтая таблица Координаты № заказа: B' + str(int(genYellowCoord[a]) - 5) + ':D' + str(int(genYellowCoord[a]) - 3))
        #genHeadPaint = gsf.get_effective_format(worksheet, 'B' + genZakazCoord[0] + ':F' + genZakazCoord[0])
        print('Желтая таблица  Координаты Шапки ген заказа: B' + genZakazCoord[0] + ':F' + genZakazCoord[0])
        gsfb.format_cell_range(worksheetYellow, 'B' + genYellowCoord[a] + ':F' + genYellowCoord[a], genHeadPaint)
        gsfb.format_cell_range(worksheetYellow, 'B' + str(int(genYellowCoord[j]) + 1) + ':F' + str(int(genYellowCoord[j]) + 3),genHeadPaint2)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(genZakazCoord[0]) + 1) + ':F' + str(int(calendarZakazCoord[0]) - 2))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(genYellowCoord[a]) + 1) + ':F' + str(int(calendarYellowCoord[a]) - 2),genDataPaint)

    b = -1
    for i in calendarYellowCoord:
        b = b + 1
        #genHeadPaint = gsf.get_effective_format(worksheet, 'B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        gsfb.format_cell_range(worksheetYellow, 'B' + calendarYellowCoord[b] + ':F' + calendarYellowCoord[b],genHeadPaint)
        #calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        print('Координаты шапки Календарного плана:' + 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(calendarYellowCoord[b]) + 2) + ':F' + str(int(calendarYellowCoord[b]) + 4),calendarHead2Paint)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 5) + ':F' + str(int(operZakazCoord[0]) - 2))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(calendarYellowCoord[b]) + 4) + ':F' + str(int(operYellowCoord[b]) - 2),calendarDataPaint)
    c = -1
    for i in operYellowCoord:
        c = c + 1
        #genHeadPaint = gsf.get_effective_format(worksheet, 'B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        gsfb.format_cell_range(worksheetYellow, 'B' + operYellowCoord[c] + ':F' + operYellowCoord[c], operHeadPaint)
        calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        print('Координаты шапки Календарного плана:' + 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(operYellowCoord[c]) + 1) + ':F' + str(int(operYellowCoord[c]) + 2),operHead2Paint)
        genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 3) + ':F' + str(int(operZakazCoord[0]) + 3))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(operYellowCoord[c]) + 3) + ':F' + str(int(operYellowCoord[c]) + 3),operDataPaint)

    d = -1
    for i in genDoneCoord:
        d = d + 1
        #zakaz = gsf.get_effective_format(worksheet, 'C' + str(int(genZakazCoord[0]) - 4))
        gsfb.format_cell_range(worksheetDone, 'B' + str(int(genDoneCoord[d]) - 5) + ':D' + str(int(genDoneCoord[d]) - 3),zakaz)
        print('Желтая таблица Координаты № заказа: B' + str(int(genDoneCoord[d]) - 5) + ':D' + str(int(genDoneCoord[d]) - 3))
        #genHeadPaint = gsf.get_effective_format(worksheet, 'B' + genZakazCoord[0] + ':F' + genZakazCoord[0])
        print('Желтая таблица  Координаты Шапки ген заказа: B' + genZakazCoord[0] + ':F' + genZakazCoord[0])
        gsfb.format_cell_range(worksheetDone, 'B' + genDoneCoord[d] + ':F' + genDoneCoord[d], genHeadPaint)
        gsfb.format_cell_range(worksheetDone, 'B' + str(int(genDoneCoord[j]) + 1) + ':F' + str(int(genDoneCoord[j]) + 3),genHeadPaint2)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(genZakazCoord[0]) + 1) + ':F' + str(int(calendarZakazCoord[0]) - 2))
        gsfb.format_cell_range(worksheetDone,'B' + str(int(genDoneCoord[d]) + 1) + ':F' + str(int(calendarDoneCoord[d]) - 2),genDataPaint)

    e = -1
    for i in calendarDoneCoord:
        e = e + 1
        #genHeadPaint = gsf.get_effective_format(worksheet, 'B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        print('Таблица выполненные координаты шапки календарного плана: B' + calendarZakazCoord[0] + ':F' +calendarZakazCoord[0])
        gsfb.format_cell_range(worksheetDone, 'B' + calendarDoneCoord[e] + ':F' + calendarDoneCoord[e], genHeadPaint)
        #calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        print('Таблица выполненные координаты шапки календарного плана:' + 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        gsfb.format_cell_range(worksheetDone,'B' + str(int(calendarDoneCoord[e]) + 2) + ':F' + str(int(calendarDoneCoord[e]) + 4),calendarHead2Paint)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 5) + ':F' + str(int(operZakazCoord[0]) - 2))
        gsfb.format_cell_range(worksheetDone,'B' + str(int(calendarDoneCoord[e]) + 4) + ':F' + str(int(operDoneCoord[e]) - 2),genDataPaint)
    f=-1
    for i in operDoneCoord:
        f=f+1
        print('Координаты Шапки календарного плана: B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        gsfb.format_cell_range(worksheetDone, 'B' + operDoneCoord[f] + ':F' + operDoneCoord[f], operHeadPaint)
        #calendarHead2Paint = gsf.get_effective_format(worksheetDone, 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        print('Координаты шапки Календарного плана:' + 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        gsfb.format_cell_range(worksheetDone,'B' + str(int(operDoneCoord[f]) + 1) + ':F' + str(int(operDoneCoord[f]) + 2),operHead2Paint)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 3) + ':F' + str(int(operZakazCoord[0]) + 3))
        gsfb.format_cell_range(worksheetDone,'B' + str(int(operDoneCoord[f]) + 3) + ':F' + str(int(operDoneCoord[f]) + 4),operDataPaint)
    gsfb.execute()

'''def parse2(list):
    j = -1

    redTable = []
    yellowTable = []
    doneTable = []
    for i in list:
        j = j + 1
        table = getTable(i)
        gen = cutGen(table)
        headGenZakaz = cutHead(table)
        #headGen = cutGenHead(headGenZakaz)
        calendar = cutCalendar(table)
        headCalendar = cutHead(calendar)
        oper = cutOPer(table)
        headOper = cutHead(oper)
        doneTable1 = []
        redTable1 = []
        yellowTable1 = []
        doneTable2 = []
        redTable2 = []
        yellowTable2 = []
        doneTable3 = []
        redTable3 = []
        yellowTable3 = []
        scoreRed = 0
        scoreYellow = 0
        scoreDone = 0

        for i in gen:
            if len(i) >= 4:
                b = i[3:]
                if b != []:
                    if validDate(i[3]):
                        a = i[4:]
                        if a != []:
                            if validDate(i[4]):
                                doneTable1.append(i)
                        else:
                            if i[3] != '-':
                                if isItLate(i[3]):
                                    redTable1.append(i)
                                else:
                                    yellowTable1.append(i)
                            else:
                                doneTable1.append(i)

        a = []
        doneTable1.append(a)
        doneTable1.append(a)
        redTable1.append(a)
        redTable1.append(a)
        yellowTable1.append(a)
        yellowTable1.append(a)

        redTableGen = headGenZakaz + redTable1 + []
        yellowTableGen = headGenZakaz + yellowTable1 + []
        doneTableGen = headGenZakaz + doneTable1 + []

        for i in calendar:
            if len(i) >= 4:
                b = i[3:]
                if b != []:
                    if validDate(i[3]):
                        a = i[4:]
                        if a != []:
                            if validDate(i[4]):
                                doneTable2.append(i)
                        else:
                            if i[3] != '-':
                                if isItLate(i[3]):
                                    redTable2.append(i)
                                else:
                                    yellowTable2.append(i)
                            else:
                                doneTable2.append(i)

        a = []
        doneTable2.append(a)
        doneTable2.append(a)
        redTable2.append(a)
        redTable2.append(a)
        yellowTable2.append(a)
        yellowTable2.append(a)

        redTableCalendar = headCalendar + redTable2 + []
        yellowTableCalendar = headCalendar + yellowTable2 + []
        doneTableCalendar = headCalendar + doneTable2 + []

        for i in oper:
            if len(i) >= 4:
                b = i[3:]
                if b != []:
                    if validDate(i[3]):
                        a = i[4:]
                        if a != []:
                            if validDate(i[4]):
                                doneTable3.append(i)
                        else:
                            if i[3] != '-':
                                if isItLate(i[3]):
                                    redTable3.append(i)
                                else:
                                    yellowTable3.append(i)
                            else:
                                doneTable3.append(i)

        redTableOper = headOper + redTable3
        yellowTableOper = headOper + yellowTable3
        doneTableOper = headOper + doneTable3

        redTable += redTableGen + [] + redTableCalendar + redTableOper
        yellowTable += yellowTableGen + yellowTableCalendar + yellowTableOper
        doneTable += doneTableGen + doneTableCalendar + doneTableOper

        print('Красная таблица:')
        for i in redTable:
            print(i)
        print('Желтая таблица')
        for i in yellowTable:
            print(i)
        print('Выполненные таблица')
        for i in doneTable:
            print(i)

    print('отправляю таблицу:')
    for i in redTable:
        print(i)
    worksheetRed.update('B2:F' + str(len(redTable) + 2), redTable)
    scoreRed += len(redTable) + 4
    worksheetYellow.update('B2:F'+str(len(yellowTable) + 2),yellowTable)
    scoreYellow+=len(yellowTable) + 3
    worksheetDone.update('B2:F'+str(len(doneTable) + 2),doneTable)
    scoreDone+=len(doneTable)+3
    

    '''

def cutGenHead(table):
    result = table[5:]
    return result
def testPaint():
    list = ['2747','2707','2774','2776','2707-01']
    worksheet = spreadsheet.worksheet('2747')

    paintZakaz = gsf.get_effective_format(worksheet,'B2:D2')
    paintHead = gsf.get_effective_format(worksheet,'B7:F7')
    paintHead2 = gsf.get_effective_format(worksheet,'B8:F8')
    paintTable = gsf.get_effective_format(worksheet,'B11:F11')
    gsfb = gsf.batch_updater(spreadsheet)
    gsfb.format_cell_range(worksheetRed,'B2:F4',paintZakaz)
    gsfb.format_cell_range(worksheetRed,'B7:F7',paintHead)
    gsfb.format_cell_range(worksheetRed,'B8:F10',paintHead2)
    gsfb.format_cell_range(worksheetRed,'B10:F15',paintHead2)
    gsfb.execute()
#worksheet = spreadsheet.worksheet('2747')
#test = worksheet.batch_get(['B2:F10', 'B11'])
#print(test)

print(spreadsheet.worksheets())
list = ['2777','2707-02','2707-01','2776','2774','2767','2761','2754','2752','2747']
parse2(list)
testOPti2(spreadsheet)


def testGet():
    gsfGet = gsf.batch_updater(spreadsheet)
    gsfGet.get_effective_format(worksheetRed, 'B2:F2')
    worksheetRed.clear()
    gsfGet.format_cell_range(worksheetRed, 'B2:F2')
    gsfGet.execute()
#test = ['2761', '2747']
#parse2(list)
#testOPti(spreadsheet)
#testGet()
