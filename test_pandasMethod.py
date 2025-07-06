import pandas as pd
import openpyxl
import glob
import os
import numpy as np

df = pd.read_excel("../excel/adaptor/0.8-JJS.xlsx", header=0, engine='openpyxl')
df.set_index(df.columns[0], inplace=True)
stu = pd.read_excel("../../pandas_excel_exercise/outputExcel/student_scores.xlsx", header=0,engine='openpyxl')
stu.set_index('ID', inplace=True)
# print(type(stu.loc[2,'Name']))

s = 'Frequency 0 to 120GHz'
start = s.find('Frequency') + len('Frequency') + 1
value = s[start:]
print(f'the value of frequency is {value}')

# print(df.columns)
# print(df.columns[0])
# print("------------------------------")
# qualified_stu = stu[stu['Score'].apply(lambda x: x>= 90)]
# print(qualified_stu)
# print("------------------------------")
# print(qualified_stu.index)
# print(df.index)
# print(df.iloc[1])
# print(df[0].loc[0])
# print(df[0].at[0])