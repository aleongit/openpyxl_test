from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# Worksheet Tables_______________________________________________________________
# Creating a table_________________________________________

# Worksheet tables are references to groups of cells.
# This makes certain operations such as styling the cells in a table easier.

wb = Workbook()
ws = wb.active

data = [
    ['Apples', 10000, 5000, 8000, 6000],
    ['Pears',   2000, 3000, 4000, 5000],
    ['Bananas', 6000, 6000, 6500, 6000],
    ['Oranges',  500,  300,  200,  700],
]

# add column headings. NB. these must be strings
ws.append(["Fruit", "2011", "2012", "2013", "2014"])
for row in data:
    ws.append(row)

tab = Table(displayName="Table1", ref="A1:E5")

# Add a default style with striped rows and banded columns
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
tab.tableStyleInfo = style

'''
Table must be added using ws.add_table() method to avoid duplicate names.
Using this method ensures table name is unque through out defined names and all other table name. 
'''
ws.add_table(tab)
wb.save("table.xlsx")

# Table names must be unique within a workbook.
# By default tables are created with a header from the first row and filters for all the columns 
# and table headers and column headings must always contain strings.

# -------------------------------------------------------------------------------------------
# Warning !!!
# In write-only mode you must add column headings to tables manually 
# and the values must always be the same as the values of the corresponding cells 
# (ee below for an example of how to do this),
# otherwise Excel may consider the file invalid and remove the table.
# -------------------------------------------------------------------------------------------

# Styles are managed using the the TableStyleInfo object.
# This allows you to stripe rows or columns and apply the different colour schemes.


# Working with Tables_________________________________________

# ws.tables is a dictionary-like object of all the tables in a particular worksheet:

ws.tables
# print(ws.tables)
# {"Table1",  <openpyxl.worksheet.table.Table object>}

# Get Table by name or range
ws.tables["Table1"]
# print(ws.tables["Table1"])

#or
# print(ws.tables["A1:E5"])

# Iterate through all tables in a worksheet
for table in ws.tables.values():
    #print(table)
    for item in table:
        print(item)

# get range to get value
rang = ws[ws.tables['Table1'].ref]
print(rang)
for row in rang:
    #print(row)
    for cel in row:
        print(cel.value, end=' ')
    print()

# Get table name and range of all tables in a worksheet
# Returns a list of table name and their ranges.
ws.tables.items()
print(ws.tables.items)
# [("Table1", "A1:D10")]

# The number of tables in a worksheet
print(len(ws.tables))
# 1

# Delete a table
del ws.tables["Table1"]


# Manually adding column headings_________________________________________

# In write-only mode you can either only add tables without headings:

table.headerRowCount = False

# Or initialise the column headings manually:
headings = ["Fruit", "2011", "2012", "2013", "2014"] # all values must be strings
table._initialise_columns()
for column, value in zip(table.tableColumns, headings):
    column.name = value