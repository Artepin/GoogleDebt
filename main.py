
import gspread                                      #Подключение библиотеки для взаимодействия с google sheets
from datetime import datetime                       #Импорт модуля сравнения даты и времени (используется только дата)
from datetime import date                           #Импорт модуля сравнения даты
import re                                           #Импорт модуля соответствия переменной регулярному выражению
import gspread_formatting as gsf                    #Импорт модуля управления форматированием таблицы
import time                                         #Импорт модуля управления временем
gp = gspread.service_account(filename='./auth.json')#Авторизация сервисного аккаунта через файл auth.json
#spreadsheet = gp.open(' РСС ведение заказов')      #Подключение таблицы РСС заказы
spreadsheet = gp.open('TestParseMyProg')            #Подключение  тестовой таблицы для проверки новых функций
worksheetRed = spreadsheet.worksheet("красные")     #Поключение листа "Красные" для дальнейшей работы в таблице
worksheetYellow = spreadsheet.worksheet("желтые")   #Поключение листа "Желтые" для дальнейшей работы в таблице
worksheetDone = spreadsheet.worksheet("выполненные")#Поключение листа "выполненные" для дальнейшей работы в таблице

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

    elif int(day)>14:
        return 2
    else:
        return False


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
            table2=table[0]
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
                for k in range(j-2, j + 1):
                    head.append(table[k])
                return head

def getList():
    worksheet = spreadsheet.worksheet('Заказы РСС')
    length = str(len(worksheet.col_values(3)))
    table2 = worksheet.batch_get(['C7:L'+length])
    table = table2[0]
    list = []
    print(table)
    for i in table:
        if i[8] =='1':
            list.append(i[0])
    print(list)
    return  list

def parse2(list):
    j=-1
    redTable = []
    yellowTable = []
    doneTable = []
    for i in list:
        #print(str(j+3)+'Таблица')
        print(list[j+1])
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
                                if isItLate(i[3])==2:
                                    continue
                                elif isItLate(i[3]):
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


        redTableGen = headGenZakaz + redTable1
        redTableGen.append(a)
        redTableGen.append(a)
        yellowTableGen = headGenZakaz + yellowTable1
        yellowTableGen.append(a)
        yellowTableGen.append(a)
        doneTableGen = headGenZakaz + doneTable1
        doneTableGen.append(a)
        doneTableGen.append(a)

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
                                if isItLate(i[3]) == 2:
                                    continue
                                elif isItLate(i[3]):
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

        redTableCalendar = headCalendar + redTable2
        redTableCalendar.append(a)
        redTableCalendar.append(a)
        yellowTableCalendar = headCalendar + yellowTable2
        yellowTableCalendar.append(a)
        yellowTableCalendar.append(a)
        doneTableCalendar = headCalendar + doneTable2
        doneTableCalendar.append(a)
        doneTableCalendar.append(a)

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
                                if isItLate(i[3]) == 2:
                                    print(i)
                                    continue
                                elif isItLate(i[3]):
                                    redTable3.append(i)
                                else:
                                    yellowTable3.append(i)
                            else:
                                doneTable3.append(i)

        redTableOper = headOper+ redTable3
        redTableOper.append(a)
        redTableOper.append(a)
        yellowTableOper = headOper+yellowTable3
        yellowTableOper.append(a)
        yellowTableOper.append(a)
        doneTableOper = headOper+doneTable3
        doneTableOper.append(a)
        doneTableOper.append(a)

        redTable += [] + redTableGen + redTableCalendar + redTableOper
        redTable.append(a)

        yellowTable += yellowTableGen + yellowTableCalendar + yellowTableOper
        yellowTable.append(a)

        doneTable += doneTableGen + doneTableCalendar + doneTableOper
        doneTable.append(a)



    print('отправляю таблицу красные:')
    for i in redTable:
        print(i)
    print('Отправляю таблицу жёлтые:')
    for i in yellowTable:
        print(i)
    print('Отправляю таблицу выполненные:')
    for i in doneTable:
        print(i)

    print(len(redTable))
    #time.sleep(30)
    worksheetRed.clear()
    worksheetRed.update('B2:F' + str(len(redTable) + 2), redTable)

    worksheetYellow.clear()
    worksheetYellow.update('B2:F'+str(len(yellowTable) + 2),yellowTable)

    worksheetDone.clear()
    worksheetDone.update('B2:F'+str(len(doneTable) + 2),doneTable)

def testOPti2(spreadsheet):
    gsfb = gsf.batch_updater(spreadsheet)
    worksheet = spreadsheet.worksheet('шаблон')
    genZakazCoord, calendarZakazCoord, operZakazCoord = test3(worksheet)
    genRedCoord, calendarRedCoord, operRedCoord = test3(worksheetRed)
    genYellowCoord, calendarYellowCoord, operYellowCoord = test3(worksheetYellow)
    genDoneCoord, calendarDoneCoord, operDoneCoord = test3(worksheetDone)

    #clear = gsf.get_effective_format(worksheetRed,'F2')
    #gsf.format_cell_range(worksheetRed,'B2:F'+ str(int(operRedCoord[0]) + 3),clear)
    zakaz = gsf.get_effective_format(worksheet, 'C' + str(int(genZakazCoord[0]) - 4))
    genHeadPaint = gsf.get_effective_format(worksheet, 'B' + genZakazCoord[0] + ':F' + genZakazCoord[0])
    genHeadPaint2 = gsf.get_effective_format(worksheet, 'B' + str(int(genZakazCoord[0]) + 1) + ':F' + str(int(genZakazCoord[0]) + 1))
    genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(genZakazCoord[0]) + 4) + ':F' + str(int(genZakazCoord[0]) + 4))

    calendarHeadPaint = gsf.get_effective_format(worksheet, 'B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
    calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 1) + ':F' + str(int(calendarZakazCoord[0]) + 3))
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
        gsfb.format_cell_range(worksheetYellow, 'B' + str(int(genYellowCoord[a]) + 1) + ':F' + str(int(genYellowCoord[a]) + 3),genHeadPaint2)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(genZakazCoord[0]) + 1) + ':F' + str(int(calendarZakazCoord[0]) - 2))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(genYellowCoord[a]) + 5) + ':F' + str(int(calendarYellowCoord[a]) - 2),genDataPaint)

    b = -1
    for i in calendarYellowCoord:
        b = b + 1
        #genHeadPaint = gsf.get_effective_format(worksheet, 'B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + calendarZakazCoord[0] + ':F' + calendarZakazCoord[0])
        gsfb.format_cell_range(worksheetYellow, 'B' + calendarYellowCoord[b] + ':F' + calendarYellowCoord[b],genHeadPaint)
        #calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        print('Координаты шапки Календарного плана:' + 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(calendarYellowCoord[b]) + 1) + ':F' + str(int(calendarYellowCoord[b]) + 4),calendarHead2Paint)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 5) + ':F' + str(int(operZakazCoord[0]) - 2))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(calendarYellowCoord[b]) + 4) + ':F' + str(int(operYellowCoord[b]) - 2),calendarDataPaint)
    c = -1
    for i in operYellowCoord:
        c = c + 1
        #genHeadPaint = gsf.get_effective_format(worksheet, 'B' + operZakazCoord[0] + ':F' + operZakazCoord[0])
        print('Координаты Шапки календарного плана: B' + operYellowCoord[c] + ':F' + operYellowCoord[c])
        gsfb.format_cell_range(worksheetYellow, 'B' + operYellowCoord[c] + ':F' + operYellowCoord[c], operHeadPaint)
        #calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        print('Координаты шапки Календарного плана:' + 'B' + str(int(operYellowCoord[c]) + 1) + ':F' + str(int(operYellowCoord[c]) + 2))
        gsfb.format_cell_range(worksheetYellow,'B' + str(int(operYellowCoord[c]) + 1) + ':F' + str(int(operYellowCoord[c]) + 2),operHead2Paint)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 3) + ':F' + str(int(operZakazCoord[0]) + 3))
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
        #print('Таблица выполненные координаты шапки календарного плана: B' + calendarZakazCoord[0] + ':F' +calendarZakazCoord[0])
        gsfb.format_cell_range(worksheetDone, 'B' + calendarDoneCoord[e] + ':F' + calendarDoneCoord[e], genHeadPaint)
        #calendarHead2Paint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        #print('Таблица выполненные координаты шапки календарного плана:' + 'B' + str(int(calendarZakazCoord[0]) + 2) + ':F' + str(int(calendarZakazCoord[0]) + 4))
        gsfb.format_cell_range(worksheetDone,'B' + str(int(calendarDoneCoord[e]) + 2) + ':F' + str(int(calendarDoneCoord[e]) + 4),calendarHead2Paint)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(calendarZakazCoord[0]) + 5) + ':F' + str(int(operZakazCoord[0]) - 2))
        gsfb.format_cell_range(worksheetDone,'B' + str(int(calendarDoneCoord[e]) + 4) + ':F' + str(int(operDoneCoord[e]) - 2),genDataPaint)
    f=-1
    for i in operDoneCoord:
        f=f+1
        #print('Координаты оперативных задач таблицы выполнено: B' + operDoneCoord[f] + ':F' + operZakazCoord[f])
        gsfb.format_cell_range(worksheetDone, 'B' + operDoneCoord[f] + ':F' + operDoneCoord[f], operHeadPaint)
        #calendarHead2Paint = gsf.get_effective_format(worksheetDone, 'B' + str(int(operZakazCoord[0]) + 1) + ':F' + str(int(operZakazCoord[0]) + 2))
        #print('Координаты оперативных задач таблицы выполнено 2: B' + str(int(operDoneCoord[f]) + 2) + ':F' + str(int(operDoneCoord[f]) + 2))
        gsfb.format_cell_range(worksheetDone,'B' + str(int(operDoneCoord[f]) + 1) + ':F' + str(int(operDoneCoord[f]) + 2),operHead2Paint)
        #genDataPaint = gsf.get_effective_format(worksheet, 'B' + str(int(operZakazCoord[0]) + 3) + ':F' + str(int(operZakazCoord[0]) + 3))
        gsfb.format_cell_range(worksheetDone,'B' + str(int(operDoneCoord[f]) + 3) + ':F' + str(int(operDoneCoord[f]) + 4),operDataPaint)
    gsfb.execute()

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

def itIsDate(data):
    print(data)
    validFull = re.search(r'\d{2}.\d{2}.\d{4}',data)
    if validFull:
        return True
    else:
        validPart = re.search(r'\d{2}.\d{2},\d{2}',data)
        if validPart:
            return True
        else:
            return False

def getRed():
    b = worksheetRed.col_values(2)
    redTable2 =  worksheetRed.batch_get(['B2:F'+str(len(b))])
    redTable = redTable2[0]
    for i in redTable:
        print(i)
    #print(len(b))
    return redTable

def parseData(table):
    a = 0
    newDates = []
    for i in table:
        a+=1
        if len(i)>4:
            if itIsDate(str(i[4])):
                newDates.append(i)
    print('Таблица с новыми датами:')
    print(newDates)
    return newDates

def findNmber(table):
    number = []
    coord = []
    j=0
    for i in table:
        if i!=[]:
            find = re.search(r'\d{4}',i[0])
            if find:
                number.append('B'+str(int(j)+2)+' '+i[0])
        j = j + 1
    print(number)


def parseRed():
    table = getRed()
    newDates =parseData(table)
    findNmber(table)




def exportListOfSheets():
    lst = []
    lst = spreadsheet.worksheets()
    #print(list)
    result = []
    for i in lst:
        cutNumber = str(i)[12:19]

        if re.search(r'\d{4}-\d{2}',cutNumber):
            print('Заказ через тире')
            print(cutNumber)
            result.append(cutNumber)
        elif re.search(r'\d{4}',cutNumber):
            number = cutNumber[:4]
            print(number)
            result.append(number)
    print(result)
    return result

list3 = ['2613','2634','2650',
        '2691','2692','2707-01', '2707-02',
        '2716','2739', '2747',
        '2752','2754','2761','2764','2767',
        '2776','2777']
warn = ['2150','2673','2686','2714','2715',]
listMy = ['2634','2777', '2707-02','2707-01', '2776','2774','2767', '2761','2754','2752','2747' ]

parseRed()
#list = getList()
#parse2(listMy)
#time.sleep(30)
#testOPti2(spreadsheet)


