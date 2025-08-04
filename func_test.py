import camelot
import pandas as pd
import openpyxl
import glob
import os
import numpy as np
import py
import ctypes


import excel_organization_func as ef

#--------------------------------------------
# df = pd.read_excel("../excel/adapter/0.8-JJS.xlsx", header=0, engine='openpyxl')
# df.set_index(df.columns[0], inplace=True)
#---------------------------------------------
# stu = pd.read_excel("../../pandas_excel_exercise/outputExcel/student_scores.xlsx", header=0,engine='openpyxl')
# stu.set_index('ID', inplace=True)
# mask = stu.astype(str).apply(lambda col: ef.str_match(col, 'student_002'))
# tar_row = mask.loc[2,]
# print(stu)
# print(tar_row[tar_row].index)
# print('done')

# def load_ghostscript():
#     try:
#         path = "/opt/homebrew/lib/libgs.dylib"
#         ctypes.CDLL(path)
#         print("Ghostscript loaded.")
#         return True
#     except OSError as e:
#         print("Failed to load Ghostscript:", e)
#         return False
# load_ghostscript()
tables = camelot.read_pdf('../../datasheets/coaxial connector_end launch/FR2-SMA-KHD23_2025-07-10.pdf', pages='all')
print(tables)
#print(tables[0].df)
print('done')



