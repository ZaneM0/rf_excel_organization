import pandas as pd
import openpyxl
import glob
import os
import numpy as np
import excel_organization_func as ef

df = pd.read_excel("../excel/adapter/0.8-JJS.xlsx", header=0, engine='openpyxl')
df.set_index(df.columns[0], inplace=True)
# stu = pd.read_excel("../../pandas_excel_exercise/outputExcel/student_scores.xlsx", header=0,engine='openpyxl')
# stu.set_index('ID', inplace=True)
# print(type(stu.loc[2,'Name']))

find = ef.str_match_bool(df,r'\s*cent(?:er|re)\s*')
print(find)
print('done!')
# s = 'Frequency 0 to 120GHz'
# start = s.find('Frequency') + len('Frequency') + 1
# value = s[start:]
# print(f'the value of frequency is {value}')
