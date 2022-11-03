from openpyxl import Workbook

# Create a workbook_______________________________________________________________

wb = Workbook()

# Workbook.active
ws = wb.active

# -------------------------------------------------------------------------------------------
# Note !!!
# This is set to 0 by default. 
# Unless you modify its value, you will always get the first worksheet by using this method.
# -------------------------------------------------------------------------------------------

# Workbook.create_sheet()
ws1 = wb.create_sheet("End") # insert at the end (default)
# or
ws2 = wb.create_sheet("First", 0) # insert at first position
# or
ws3 = wb.create_sheet("Penultimate", -1) # insert at the penultimate position

# Title
ws.title = "New"

# Tab color HEX
#ws.sheet_properties.tabColor = "1072BA"
ws.sheet_properties.tabColor = "E2FFA7"

# Get worksheet
ws = wb["New"]

# Sheetnames
print(wb.sheetnames)
#['Sheet2', 'New Title', 'Sheet1']

# Sheet Loop
for sheet in wb:
    print(sheet.title)

# Set active sheet
# wb.active = wb['sheet_name']
wb.active = wb['New']
print(wb.active)

# Copy worksheet
source = wb.active
target = wb.copy_worksheet(source)

# -------------------------------------------------------------------------------------------
# Note !!!
# Only cells (including values, styles, hyperlinks and comments) 
# and certain worksheet attribues (including dimensions, format and properties) are copied. 
# All other workbook / worksheet attributes are not copied - e.g. Images, Charts.
# You also cannot copy worksheets between workbooks. 
# You cannot copy a worksheet if the workbook is open in read-only or write-only mode.
# -------------------------------------------------------------------------------------------

# Playing with data_______________________________________________________________
# Accessing one cell_________________________________________

# Get cell
c = ws['A4']
print('cell', c)

# Set cell
ws['A1'] = 4
ws['A2'] = 'test'

# There is also the Worksheet.cell() method.
# This provides access to cells using row and column notation:
d = ws.cell(row=4, column=2, value=10)

# -------------------------------------------------------------------------------------------
# Note !!!
# When a worksheet is created in memory, it contains no cells.
# They are created when first accessed.
# 
# Warning !!!
# Because of this feature, scrolling through cells 
# instead of accessing them directly will create them all in memory, 
# even if you don’t assign them a value.
#
# Something like
#
# >>> for x in range(1,101):
# ...        for y in range(1,101):
# ...            ws.cell(row=x, column=y)
#
# will create 100x100 cells in memory, for nothing.
# -------------------------------------------------------------------------------------------

# Save the file
wb.save("tutorial.xlsx")