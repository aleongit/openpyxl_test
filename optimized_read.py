from openpyxl import load_workbook
import time

# Optimised Modes_______________________________________________________________
# Read-only mode_________________________________________

# Sometimes, you will need to open or write extremely large XLSX files, 
# and the common routines in openpyxl won’t be able to handle that load.
# Fortunately, there are two modes that enable you to read and 
# write unlimited amounts of data with (near) constant memory consumption.

#Introducing openpyxl.worksheet._read_only.ReadOnlyWorksheet:

# get the start time
st = time.time()

wb = load_workbook(filename='lf.xlsx', read_only=True)
ws = wb['big_data']

for row in ws.rows:
    for cell in row:
        print(cell.value, end=' ')
    print()

# Close the workbook after reading
wb.close()

# get the end time
et = time.time()

# get the execution time
elapsed_time = round(et - st, 2)
print('Execution time read_only:', elapsed_time, 'seconds')

# -------------------------------------------------------------------------------------------
# Warning !!!
# . openpyxl.worksheet._read_only.ReadOnlyWorksheet is read-only
# . Unlike a normal workbook, a read-only workbook will use lazy loading. 
# The workbook must be explicitly closed with the close() method.
# -------------------------------------------------------------------------------------------

# Cells returned are not regular 
#   openpyxl.cell.cell.Cell 
# but 
#   openpyxl.cell._read_only.ReadOnlyCell.

# Worksheet dimensions
# Read-only mode relies on applications and libraries that created 
# the file providing correct information about the worksheets, 
# specifically the used part of it, known as the dimensions. 
# Some applications set this incorrectly. 
# You can check the apparent dimensions of a worksheet using ws.calculate_dimension(). 
# If this returns a range that you know is incorrect, 
# say A1:A1 then simply resetting the max_row and max_column attributes 
# should allow you to work with the file:

# ws.reset_dimensions()