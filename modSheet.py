import os
import xlwings as xw
import psutil
from datetime import datetime, timedelta
import win32com.client

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

def modifySheets(folder, month, year, statusLabel, processStop):
    app = xw.App(visible=False)
    excelFiles = [file for file in os.listdir(folder) if file.endswith('.xlsx') and not file.startswith('~$')]
    totalFiles, filesWritten = len(excelFiles), 0

    closeExcelFiles(excelFiles)

    for fileName in excelFiles:
        if processStop.value:
            processStop.value = False
            print("Processed stopped")
            break
        try:
            filePath = os.path.join(folder, fileName)
            wb = xw.Book(filePath, ignore_read_only_recommended=True)
            ws = wb.sheets[0]
            currentMonth = ws.range('F8').value

            # New Template
            if currentMonth and currentMonth.capitalize() in months:
                ws.range('F8').value = month.upper()
                ws.range('H8').value = year

            # Old Template
            else:
                ws.range('A12:B40').clear_contents()
                ws.range('A12:H40').api.Borders.LineStyle = None

                days = ws.range('C8').value.split()
                data = getWeekdays(month, year, days)

                ws.range('A12').value = data
                ws.range(f'A12:H{12 + len(data) - 1}').api.Borders.LineStyle = 1
            wb.save(filePath)
            wb.close()
            filesWritten += 1
            statusLabel.configure(text=f"Files written: {filesWritten}/{totalFiles}")
            statusLabel.update()
            print(f"File '{fileName}' written.")
        except FileNotFoundError:
            print(f"File '{fileName}' not found.")

    app.quit()

def printSheets(folder, statusLabel, processStop):
    excelFiles = [file for file in os.listdir(folder) if file.endswith('.xlsx') and not file.startswith('~$')]
    totalFiles, filesPrinted = len(excelFiles), 0

    closeExcelFiles(excelFiles)

    for fileName in excelFiles:
        if processStop.value:
            processStop.value = False
            print("Processed stopped")
            break
        try:
            filePath = os.path.join(folder, fileName)
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            workbook = excel.Workbooks.Open(filePath)
            workbook = workbook.Worksheets(1)
            workbook.PrintOut()
            excel.Quit()
            
            filesPrinted += 1
            statusLabel.configure(text=f"Files printed: {filesPrinted}/{totalFiles}")
            statusLabel.update()
            print(f"File '{fileName}' printed.")
        except FileNotFoundError:
            print(f"File '{fileName}' not found.")

def getWeekdays(month, year, days):
    year, month = int(year), months[month]
    numDays = (datetime(year, month % 12 + 1, 1) - timedelta(days=1)).day
    monthDays = []

    weekdayIntegers = set()
    for day in days:
        weekdayIntegers.add(weekDays[day.upper()])

    for day in range(1, numDays+1):
        date = datetime(year, month, day)
        if date.weekday() in weekdayIntegers:
            monthDays.append([chineseWeekday[date.weekday()], date.strftime("%m/%d/%Y")])

    return monthDays

def closeExcelFiles(excelFiles):
    for proc in psutil.process_iter():
        try:
            if 'EXCEL.EXE' in proc.name():
                for item in proc.open_files():
                    if any(file.lower() in item.path.lower() for file in excelFiles):
                        proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

# def resizeRows(folder, statusLabel, processStop):
#     app = xw.App(visible=False)
#     excelFiles = [file for file in os.listdir(folder) if file.endswith('.xlsx') and not file.startswith('~$')]
#     totalFiles, filesWritten = len(excelFiles), 0

#     closeExcelFiles(excelFiles)

#     for fileName in excelFiles:
#         if processStop.value:
#             processStop.value = False
#             print("Processed stopped")
#             break
#         try:
#             filePath = os.path.join(folder, fileName)
#             wb = xw.Book(filePath, ignore_read_only_recommended=True)
#             ws = wb.sheets[0]

#             ws.range('4:8').row_height = 18
#             ws.range('12:34').row_height = 25
#             ws.api.ResetAllPageBreaks() 
#             ws.api.HPageBreaks.Add(Before=ws.range('A35').api)

#             ws = wb.sheets[1]
#             ws.range('J21').value = None
#             ws.range('K21').value = None

#             formula = '=FILTER($H$2:$H$32, (ROW($H$2:$H$32))*($H$2:$H$32<>""))'
#             ws.range('J2').formula2 = formula
#             ws.range('K2').formula2 = formula

#             wb.save(filePath)
#             wb.close()
#             filesWritten += 1
#             statusLabel.configure(text=f"Files written: {filesWritten}/{totalFiles}")
#             statusLabel.update()
#             print(f"File '{fileName}' written.")
#         except FileNotFoundError:
#             print(f"File '{fileName}' not found.")

#     app.quit()