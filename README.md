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