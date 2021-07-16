import gspread
import datetime
import re
gp = gspread.service_account(filename='./auth.json')
spreadsheet = gp.open('TestParsing')
worksheet = spreadsheet.get_worksheet(0)
dateNow = datetime.date.today()
print(dateNow)
yacheayka = 'D34'
datePlan = worksheet.acell(yacheayka).value
dateFact = worksheet.acell("E32").value
if dateFact == None:
    worksheet.format("E32", {
                "backgroundColor": {
                    "red": 1.0,
                    "green": 0.0,
                    "blue": 0.0
                }
            }
                        )


print(datePlan)
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
    match =  re.search(r'\d{2}',data)
    print ("true" if match else 'false')


def changeOfColor(coord):
    worksheet.format(coord, {
                "backgroundColor": {
                    "red": 1.0,
                    "green": 0.0,
                    "blue": 0.0
                }
            }
                        )

def prohod(dataColumn):
    for i in dataColumn:
        match = prov(dataColumn[i])
        if match:
            print("Match true")
            cellCoord = 'E'+i
            cell = worksheet.acell(cellCoord).value

            if prov(cell):
                print("change color")
                changeOfColor(cellCoord)
            else:
                print("No date")
        else:
            print("Match False")


column = worksheet.col_values(4)
prohod(column)


datePlanForm = dateTransform(datePlan)
countDays = dateRazn(datePlanForm, dateNow)
result = redOrYellow(countDays)
print(datePlanForm)
print(countDays)
print(result)
print(col)
print(prov(a1))