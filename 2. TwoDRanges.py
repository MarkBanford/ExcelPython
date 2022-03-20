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
sheet.range('A9').value = [['abc', 1, 2, 3], ['efg', 123, None, None]]  # must use None to ensure same size

# expanding

print(sheet.range('A9').expand('table'))  # table is default
time.sleep(1)
sheet.range('A9').expand().clear_contents()
