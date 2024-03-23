import os
import xlwings as xw

months = ["January", 
          "February", 
          "March", 
          "April", 
          "May", 
          "June", 
          "July", 
          "August", 
          "September", 
          "October", 
          "November", 
          "December"]

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
            weekdays = ws.range('C8').value.split()
            print(weekdays)

        wb.save(file_name)
        wb.close()

    app.quit()