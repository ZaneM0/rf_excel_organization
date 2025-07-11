import pandas as pd
import openpyxl
import glob
import os
import re
import numpy as np

def str_match(col: pd.Series, target_str: str)->pd.Series:
    return col.str.contains(target_str,case=False,na=False, regex=True)

def str_match_bool(df: pd.DataFrame, target_str: str):
    mask = df.astype(str).apply(lambda col: str_match(col, target_str))
    tar_row_mask = mask.any(axis=1)
    find = tar_row_mask.any()
    return find

def str_loc(file: str,df: pd.DataFrame, target_str: str, tar_index: int):
    mask = df.astype(str).apply(lambda col: str_match(col, target_str))
    tar_rows_mask = mask.any(axis=1)
    if not tar_rows_mask.any():
        print(f'Unable to find <{target_str}> in <{file}>! :(')
        tar_row_index = -1
        tar_col_index = -1
    else:
        if tar_index < len(tar_rows_mask[tar_rows_mask].index):
            tar_row_index = tar_rows_mask[tar_rows_mask].index[tar_index]
            # tar_cols_mask = mask.any(axis=0)
            # tar_col_index = tar_cols_mask[tar_cols_mask].index[0]
            tar_cols_mask = mask.loc[tar_row_index,]
            tar_col_index = tar_cols_mask[tar_cols_mask].index[0]
        else:
            tar_row_index = -1
            tar_col_index = -1
    return tar_row_index, tar_col_index

def get_value_same_row(file: str,df: pd.DataFrame, target_str: str, tar_index: int):
    tar_row_index, tar_col_index = str_loc(file, df, target_str, tar_index)
    value = "N/A"
    if tar_row_index == -1 and tar_col_index == -1:
        value = "N/A"
    else:
        for col in df.columns[tar_col_index + 1:]:
            if not pd.isna(df.loc[tar_row_index, col]):
                value = df.loc[tar_row_index, col]
                break
        if value == "N/A":
            for col in df.columns[:tar_col_index]:
                if not pd.isna(df.loc[tar_row_index, col]):
                    value = df.loc[tar_row_index, col]
                    break

    return value

def get_value_same_unit(file: str,df: pd.DataFrame, target_str: str, tar_index: int):
    tar_row_index, tar_col_index = str_loc(file,df, target_str, tar_index)
    if tar_row_index == -1 and tar_col_index == -1:
        value = "N/A"
    else:
        tar_unit = df.loc[tar_row_index,tar_col_index]
        value = re.sub(target_str,'',tar_unit)
        if value == '':
            value = "N/A"
        # value_start_tail = tar_unit.find(target_str) + len(target_str) + 1
        # value_start_head = tar_unit.find(target_str) - 1
        # if len(tar_unit) == len(target_str):
        #     #print(f'Unable to find the value of <{target_str}> at the same unit in <{file}>! :(')
        #     value = "N/A"
        # elif tar_unit[0] == target_str[0]:
        #     value = tar_unit[value_start_tail:]
        # else:
        #     value = tar_unit[:value_start_head]
    return value

def get_value(file: str,df: pd.DataFrame, target_str: str, tar_index: int):
    value = get_value_same_row(file, df, target_str, tar_index)
    if value == "N/A":
        value = get_value_same_unit(file, df, target_str, tar_index)
    if value == "N/A":
        print(f'Unable to find the value of <{target_str}> in <{file}>! :(>')
    return value

def get_product_name(file: str,df: pd.DataFrame):
    name_row = df.iloc[1]
    non_null = name_row.dropna().values
    if not non_null[1]:
        print(f'Unable to find product name in <{file}>! :(')
        name = "N/A"
    else:
        product_name = non_null[1]
    return product_name

def extract_from_file(file:str, parameters_dict: dict):
    df = pd.read_excel(file, header=0, engine='openpyxl')
    df.set_index(df.columns[0], inplace=True)
    param_counter = {param_name_regx: 0 for param_name_regx in parameters_dict.values()}
    param_values = []
    product_name = get_product_name(file,df)
    for param_name_regx in parameters_dict.values():
        param_values.append(get_value(file, df, param_name_regx, param_counter[param_name_regx]))
        param_counter[param_name_regx] += 1

    return product_name, param_values


def extract_from_folder(path: str, parameters_dict: dict):
    # path = "../excel/adaptor"
    files = glob.glob(os.path.join(path, '*.xlsx'))
    # param_names = ['Connector 1 Type', 'Connector 1 Impedance', 'Connector 1 Polarity',
    #                'Connector 2 Type', 'Connector 2 Impedance', 'Connector 2 Polarity',
    #                'Connector Mount Method', 'Adapter Body Style', 'Frequency', 'Insertion Loss (dB)',
    #                'VSWR /Return Loss', 'Center Contact', 'Outer Contact', 'Body', 'Dielectric',
    #                'Temperature Range', 'Compliant']
    product_info_rows = []
    param_names = list(parameters_dict.keys())
    for file in files:
        product_name, param_values = extract_from_file(file, parameters_dict)
        product_info_rows.append([product_name] + param_values)

    result = pd.DataFrame(product_info_rows, columns=['Product Name'] + param_names)
    result.set_index('Product Name', inplace=True)
    # for i in result.index:
    #     insertion_loss = result["Insertion Loss (dB)"].at[i]
    #     if 'sqt' in insertion_loss.lower():
    #         result['Insertion Loss (dB)'].at[i] = '≤' + insertion_loss[2:]

    return result

# result.to_excel('../excel/Adaptor_combined_result.xlsx', index=True, sheet_name='Adaptor')
# print(f'combined_result.xlsx has been generated successfully，include {len(product_info_rows)} products.')

# if __name__ == '__main__':
#     main()