import xlwings as xw
import numpy as np
import pandas as pd
import datetime as dt
import time
import sys
import os

data = np.random.rand(100, 100)

# Connect to Book
wb1 = xw.Book()
time.sleep(2)
xw.view(data, wb1.sheets.active)

# Sheet Object
sheet = wb1.sheets[0]
time.sleep(1)
sheet.clear()

# The Range object
sheet.range('A1').value = 10  # write
print(sheet.range('A1').value)  # READ

sheet.range('A3:B4').value = 123

# Dates
sheet.range('A6').value = dt.datetime(2021, 12, 9, 12, 3, 25)

# formula
sheet.range('B6').formula = '=SUM(B1:B4)'

# Named Ranges
sheet.range('B1').name = 'test'
