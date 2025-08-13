import camelot
import pandas as pd
import openpyxl
import glob
import os
import numpy as np
import py
import ctypes


import excel_organization_func as ef


tables = camelot.read_pdf('../../datasheets/BNC-SMB-KYK.pdf', pages='all')
# print(tables)
# print(tables[0].df)
output_excel = tables[0].df
output_df = pd.DataFrame(output_excel)
output_df.to_excel('../../datasheets/BNC-SMB-KYK.xlsx')

print('done')



