import pandas as pd
import openpyxl
import glob
import os
import numpy as np
import shutil
import re
import excel_organization_func as ef
from excel_organization_func import get_product_name

categories = ['connector','adapter','cable assembly', 'unsorted']
for cat in categories:
    os.makedirs(f'../excel/{cat}', exist_ok=True)
copied = {cat: set() for cat in categories}

folder_path = '../excel/0_excel'
files = glob.glob(os.path.join(folder_path, '*.xlsx'))

for file in files:
    assigned = False
    file_name = os.path.basename(file)
    stem = os.path.splitext(file_name)[0] #remove.xlsx
    base = re.sub(r'(?:\s*[(（]\s*\d+[)）])+$', '', stem)

    all_sheets = pd.read_excel(file, header=0, sheet_name=None,engine='openpyxl')
    for df in all_sheets.values():
        df.set_index(df.columns[0], inplace=True)
        if ef.str_match_bool(df,r'\s*connect(?:er|or)\s*datasheet'):
            assigned = True
            if get_product_name(file, df) not in copied['connector']:
                dst_folder = '../excel/connector'
                dst_path = os.path.join(dst_folder, file_name)
                shutil.copy(file, dst_path)
                copied['connector'].add(get_product_name(file, df))
            break
        elif ef.str_match_bool(df,r'\s*adapt(?:er|or)\s*datasheet'):
            assigned = True
            if get_product_name(file,df) not in copied['adapter']:
                dst_folder = '../excel/adapter'
                dst_path = os.path.join(dst_folder, file_name)
                shutil.copy(file, dst_path)
                copied['adapter'].add(get_product_name(file,df))
            break
        elif ef.str_match_bool(df,r'\s*cable assembly\s*datasheet'):
            assigned = True
            if get_product_name(file, df) not in copied['cable assembly']:
                dst_folder = '../excel/cable assembly'
                dst_path = os.path.join(dst_folder, file_name)
                shutil.copy(file, dst_path)
                copied['cable assembly'].add(get_product_name(file,df))
            break
        else:
            continue

    if not assigned:
        if base not in copied['unsorted']:
            dst_folder = '../excel/unsorted'
            dst_path = os.path.join(dst_folder, file_name)
            shutil.copy(file, dst_path)
            copied['unsorted'].add(base)

print('done!')
