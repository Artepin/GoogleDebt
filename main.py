import gspread
import datetime
import re
gp = gspread.service_account(filename='./auth.json')
spreadsheet = gp.open('TestParsing')
worksheet = spreadsheet.get_worksheet(0)
column = worksheet.col_values(4)
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

def prov(data):
    if data == None:
        data = '0'
    match =re.search(r'\d{2}',data)
    if match:
        print("date is valid")
        return True
    else:
        print("have no date")
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
        match = re.search(r'\d[2]',i)
        if match:
            print("Match true")
            cellCoord = 'E'+str(j)
            cell = worksheet.acell(cellCoord).value
            print(cell)
            if prov(cell):
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


prohod(column)
