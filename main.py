import xlwings as xw

months = ('January', 'February', 'March', 'April', 'May', 'June', 'July',
              'August', 'September', 'October', 'November', 'December')

app = xw.App(visible=False)

# Open the workbook
file_name = 'test - Copy.xlsx'
wb = xw.Book(file_name)
ws = wb.sheets[0]

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

# Write month and year to worksheet
ws.range('F8').value = month
ws.range('H8').value = year

# Save and close the workbook
wb.save(file_name)
wb.close()

app.quit()