import camelot
import pandas as pd
import openpyxl
import glob
import os
import numpy as np
import py
import ctypes


import excel_organization_func as ef


tables = camelot.read_pdf('../../datasheets/2.92-1.0-KJS-9.pdf', pages='all')
# print(tables)
# print(tables[0].df)
output_excel = tables[0].df
output_df = pd.DataFrame(output_excel)
output_df.to_excel('../../datasheets/2.92-1.0-KJS-9.xlsx')

print('done')



