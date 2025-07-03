import pandas as pd
import openpyxl
import glob
import os
import numpy as np

df = pd.read_excel("../excel/adaptor/0.8-JJS.xlsx", header=0, engine='openpyxl')
df.set_index(df.columns[0], inplace=True)
row3 = df.iloc[1]
#row3_value = row3.values
# print(row3)
# print("----------------------")
# print(row3_value)
# print(type(row3_value))

def str_match(col: pd.Series)->pd.Series:
    return col.str.contains('Configuration',case=False,na=False)


non_null = row3.dropna().values
product_name = non_null[1]

mask = df.astype(str).apply(str_match)
config_rows = mask.any(axis=1)
print(config_rows)
config_row = config_rows[config_rows].index[0]
print(config_row)


#print(df[1])
print("done!")