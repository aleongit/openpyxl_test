import datetime
from openpyxl import Workbook, load_workbook

# Simple usage_______________________________________________________________
# Using number formats_________________________________________

# file
file = 'usage.xlsx'

# open file
wb = load_workbook(filename = file)

# new sheet at the end
sheet = wb.create_sheet(title="Number_Format")

# set date using a Python datetime
sheet['A1'] = datetime.datetime(2010, 7, 21)
sheet['A1'].number_format
# 'yyyy-mm-dd h:mm:ss'

print(sheet['A1'].value)

# save
wb.save(filename = file)