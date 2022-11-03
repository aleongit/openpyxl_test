from openpyxl import Workbook

# Create a workbook_______________________________________________________________

wb = Workbook()

# Workbook.active
ws = wb.active

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

# Access with key
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

# Playing with data_______________________________________________________________
# Accessing one cell_________________________________________


# Save the file
wb.save("tutorial.xlsx")