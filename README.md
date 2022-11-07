# openpyxl

A Python library to read/write Excel 2010 xlsx/xlsm files

- https://openpyxl.readthedocs.io/en/stable/
- Version: 3.0.10

## Introduction

openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.

It was born from lack of existing library to read/write natively from Python the Office Open XML format.

## Requeriments

- python >= 3.6
- Excel 2010 xlsx/xlsm/xltx/xltm files

## Installation

```bash
pip install openpyxl
```

## Optional

```bash
pip install pillow
```

## Sample code:

```py
from openpyxl import Workbook
wb = Workbook()

# grab the active worksheet
ws = wb.active

# Data can be assigned directly to cells
ws['A1'] = 42

# Rows can also be appended
ws.append([1, 2, 3])

# Python types will automatically be converted
import datetime
ws['A2'] = datetime.datetime.now()

# Save the file
wb.save("sample.xlsx")
```

## <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html">Tutorial</a>

```
py: [tutorial.py]
xlsx: [tutorial.xlsx]
```

1. <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html#create-a-workbook">Create a Workbook</a>

2. <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html#playing-with-data">Playing with data</a>
- <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html#accessing-one-cell">Accessing one cell</a>
- <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html#accessing-many-cells">Accessing many cells</a>
- <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html#values-only">Values only</a>

3. <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html#data-storage">Data storage</a>
- <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html#saving-to-a-file">Saving to a file</a>
- <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html#saving-as-a-stream">Saving as a stream</a>

4. <a href="https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file">Loading from a file</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/usage.html">Simple usage</a>

```
py: [
    usage_write.py,
    usage_read.py,
    usage_number.py,
    usage_formulae.py,
    usage_merge.py,
    usage_image.py,
    usage_fold.py
    ]
xlsx: [usage.xlsx]
```

1. <a href="https://openpyxl.readthedocs.io/en/stable/usage.html#write-a-workbook">Write a workbook</a>
2. <a href="https://openpyxl.readthedocs.io/en/stable/usage.html#read-an-existing-workbook">Read an existing workbook</a>
3. <a href="https://openpyxl.readthedocs.io/en/stable/usage.html#using-number-formats">Using number formats</a>
4. <a href="https://openpyxl.readthedocs.io/en/stable/usage.html#using-formulae">Using formulae</a>
5. <a href="https://openpyxl.readthedocs.io/en/stable/usage.html#merge-unmerge-cells">Merge / Unmerge cells</a>
6. <a href="https://openpyxl.readthedocs.io/en/stable/usage.html#inserting-an-image">Inserting an image</a>
7. <a href="https://openpyxl.readthedocs.io/en/stable/usage.html#fold-outline">Fold (outline)</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/performance.html">Performance</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/optimized.html">Optimised Modes</a>

```
py: [
    optimized_read.py,
    optimized_write.py
    ]
xlsx: [
    lf.xlsx (to read large file 17.890 rows),
    big_file.xlsx (to write),
    write_only_file.xlsx (to write)
    ]
```

1. <a href="https://openpyxl.readthedocs.io/en/stable/optimized.html#read-only-mode">Read-only mode</a>
2. <a href="https://openpyxl.readthedocs.io/en/stable/optimized.html#write-only-mode">Write-only mode</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html">Editing Worksheets</a>

. Inserting and deleting rows and columns, moving ranges of cells

```
py: [editing.py]
xlsx: [editing.xlsx]
```

1. <a href="https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html#inserting-rows-and-columns">Inserting rows and columns</a>
2. <a href="https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html#deleting-rows-and-columns">Deleting rows and columns</a>
3. <a href="https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html#moving-ranges-of-cells">Moving ranges of cells</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/pandas.html">Working with Pandas and NumPy</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/charts/introduction.html">Charts</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/comments.html">Comments</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/styles.html">Working with styles</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/worksheet_properties.html">Additional Worksheet Properties</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/formatting.html">Conditional Formatting</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/pivot.html">Pivot Tables</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/print_settings.html">Print Settings</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/filters.html">Using filters and sorts</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/validation.html">Validating cells</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/defined_names.html">Defined Names</a>

```
py: [defined_names.py]
xlsx: [defined_names.xlsx]
```

1. <a href="https://openpyxl.readthedocs.io/en/stable/defined_names.html#sample-use-for-ranges">Sample use for ranges</a>
2. <a href="https://openpyxl.readthedocs.io/en/stable/defined_names.html#creating-new-named-ranges">Creating new named ranges</a>


## <a href="https://openpyxl.readthedocs.io/en/stable/worksheet_tables.html">Worksheet Tables</a>

```
py: [tables.py]
xlsx: [tables.xlsx]
```

1. <a href="https://openpyxl.readthedocs.io/en/stable/worksheet_tables.html#creating-a-table">Creating a table</a>
2. <a href="https://openpyxl.readthedocs.io/en/stable/worksheet_tables.html#working-with-tables">Working with Tables</a>
3. <a href="https://openpyxl.readthedocs.io/en/stable/worksheet_tables.html#manually-adding-column-headings">Manually adding column headings</a>
