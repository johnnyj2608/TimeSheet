import os
import xlwings as xw

months = ('January', 'February', 'March', 'April', 'May', 'June', 'July',
              'August', 'September', 'October', 'November', 'December')

# Input month
while True:
    month = input("Enter a month: ").capitalize()
    if month in months:
        break
    print("Not a valid month.")

# Input year
while True:
    year = input("Enter a year: ")
    if year.isdigit():
        year = int(year)
        break
    print("Not a valid year.")

app = xw.App(visible=False)

directory = os.getcwd()

excel_files = [file for file in os.listdir(directory) if file.endswith('.xlsx')]

print(excel_files)

for file_name in excel_files:
    # Open the workbook
    wb = xw.Book(file_name)
    ws = wb.sheets[0]

    # Write month and year to worksheet
    ws.range('F8').value = month
    ws.range('H8').value = year

    # Save and close the workbook   
    wb.save(file_name)
    wb.close()

app.quit()