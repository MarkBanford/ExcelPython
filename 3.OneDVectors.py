import xlwings as xw
import numpy as np
import pandas as pd
import datetime as dt
import time
import sys
import os

wb1 = xw.Book()
time.sleep(1)
sheet = wb1.sheets.active

v = [1, 2, 3, 4]

sheet.range('A12').value = v

# Transpose

sheet.range('G1').options(transpose=True).value = v

print(sheet.range('G1').expand('down').value)
