import gspread
import datetime
gp = gspread.service_account(filename='./auth.json')
worksheet = gp.open('TestParsing')
sheet = worksheet.get_worksheet(0)

dateNow = datetime.date.today()
print(dateNow)
datePlan = sheet.acell("D32").value
dateFact = sheet.acell("E32").value
if dateFact == None:
    sheet.format("E32", {
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
        print( "Red color")
    else:
        print("Yellow color")
datePlanForm = dateTransform(datePlan)
countDays = dateRazn(datePlanForm, dateNow)
result = redOrYellow(countDays)
#dateRazn = datePlanForm-dateNow
print(datePlanForm)
print(countDays)
print(result)