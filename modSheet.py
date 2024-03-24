import os
import xlwings as xw
from datetime import datetime, timedelta

months = {"January":1, 
          "February":2, 
          "March":3, 
          "April":4, 
          "May":5, 
          "June":6, 
          "July":7, 
          "August":8, 
          "September":9, 
          "October":10, 
          "November":11, 
          "December":12
          }

weekDays = {"MON":0,
            "TUE":1,
            "WED":2,
            "THU":3,
            "FRI":4,
            "SAT":5
            }

chineseWeekday = {
    0: "星期一",
    1: "星期二",
    2: "星期三",
    3: "星期四",
    4: "星期五",
    5: "星期六",
}

def modifySheets(folder, month, year):
    app = xw.App(visible=False)
    excelFiles = [file for file in os.listdir(folder) if file.endswith('.xlsx')]

    for file_name in excelFiles:
        wb = xw.Book(file_name)
        ws = wb.sheets[0]

        # New Template
        if ws.range('F8').value in months:
            ws.range('F8').value = month
            ws.range('H8').value = year

        # Old Template
        else:
            ws.range('A12:B40').clear_contents()

            days = ws.range('C8').value.split()
            print(month, year, days)
            monthDays = getWeekdays(month, year, days)

            for i in monthDays:
                date, chinese = i, monthDays[i]
                print(chinese, date)

        wb.save(file_name)
        wb.close()

    app.quit()

def getWeekdays(month, year, days):
    year, month = int(year), months[month]
    numDays = (datetime(year, month % 12 + 1, 1) - timedelta(days=1)).day
    monthDays = {}

    weekdayIntegers = set()
    for day in days:
        weekdayIntegers.add(weekDays[day.upper()])

    for day in range(1, numDays+1):
        date = datetime(year, month, day)
        if date.weekday() in weekdayIntegers:
            monthDays[date.strftime("%m/%d/%Y")] = chineseWeekday[date.weekday()]

    return monthDays