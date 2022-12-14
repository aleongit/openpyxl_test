from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Editing Worksheets_______________________________________________________________
# Inserting rows and columns_________________________________________

# ini
file = 'editing.xlsx'
wb = Workbook()
sheet = wb.worksheets[0]
sheet.title = 'editing'

# You can insert rows or columns using the relevant worksheet methods:
# . openpyxl.worksheet.worksheet.Worksheet.insert_rows()
# . openpyxl.worksheet.worksheet.Worksheet.insert_cols()
# . openpyxl.worksheet.worksheet.Worksheet.delete_rows()
# . openpyxl.worksheet.worksheet.Worksheet.delete_cols()

# The default is one row or column.
# For example to insert a row at 7 (before the existing row 7):
sheet.insert_rows(7)


# Deleting rows and columns_________________________________________

# To delete the columns F:H:
sheet.delete_cols(6, 3)

# -------------------------------------------------------------------------------------------
# Note !!!
# Openpyxl does not manage dependencies, such as formulae, tables, charts, etc., 
# when rows or columns are inserted or deleted. 
# This is considered to be out of scope for a library that focuses on managing the file format. 
# As a result, client code must implement the functionality required in any particular use case.
# -------------------------------------------------------------------------------------------


# Moving ranges of cells_________________________________________

# You can also move ranges of cells within a worksheet:

sheet.move_range("D4:F10", rows=-1, cols=2)

# This will move the cells in the range D4:F10 up one row, and right two columns.
# The cells will overwrite any existing cells.

# If cells contain formulae you can let openpyxl translate these for you, 
# but as this is not always what you want it is disabled by default. 
# Also only the formulae in the cells themselves will be translated. 
# References to the cells from other cells or defined names will not be updated; 
# you can use the Parsing Formulas translator to do this:

sheet.move_range("G4:H10", rows=1, cols=1, translate=True)
#This will move the relative references in formulae in the range by one row and one column.

# save
wb.save(filename = file)