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