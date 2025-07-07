import pandas as pd
import openpyxl
import glob
import os

def extract_from_file(filepath):
    # 以无表头模式读入整个表格
    df = pd.read_excel(filepath, header=None, engine='openpyxl')

    # ——1. 抽取产品名（第三行，第2个非空）
    row3 = df.iloc[2]                   # 第3行，0-based index=2
    non_null = row3.dropna().values     # 丢掉空值
    if len(non_null) < 2:
        product_name = non_null[0] if len(non_null) == 1 else ''
        # equivalent to a = (condition)? b : c in C langua
    else:
        product_name = non_null[1]      # 第二个非空

    # ——2. 找到“Configuration”行和列
    # 全表格查找包含“Configuration”的单元格
    mask = df.astype(str).apply(lambda col: col.str.contains('Configuration', case=False, na=False))
    # mask has the same dataframe as df but all its elements are bool value
    config_rows = mask.any(axis=1)
    #.any(): used to check if there is at least 1 Ture in specific axis; axis = 1(row)/0(col)
    if not config_rows.any():
        raise ValueError(f'[{os.path.basename(filepath)}] unable to find Configuration title')
    config_row = config_rows[config_rows].index[0]

    config_cols = mask.loc[config_row]
    config_col = config_cols[config_cols].index[0]

    # ——3. 在 config_row+1 那一行里，往右找第一个非空列，作为值列
    next_row = df.iloc[config_row + 1]
    value_col = None
    for col in df.columns[df.columns > config_col]:
        if not pd.isna(next_row[col]):
            value_col = col
            break
    if value_col is None:
        raise ValueError(f'[{os.path.basename(filepath)}] unable to find column value under Configuration')

    # ——4. 从 config_row+1 往下遍历，直到参数名空行为止
    param_names = []
    param_values = []
    for r in range(config_row + 1, len(df)):
        name = df.iat[r, config_col]
        if pd.isna(name) or str(name).strip() == '':
            break
        param_names.append(str(name).strip())
        param_values.append(df.iat[r, value_col])

    return product_name, param_names, param_values


def main():
    # ① 修改为你所有 Excel 文件所在的文件夹
    folder = '../excel/adaptor'
    files = glob.glob(os.path.join(folder, '*.xlsx'))

    all_rows = []
    header = None

    for f in files:
        prod, names, vals = extract_from_file(f)

        # 首个文件确定表头
        if header is None:
            header = ['Product Name'] + names

        all_rows.append([prod] + vals)

    # 生成 DataFrame 并写入新文件
    result = pd.DataFrame(all_rows, columns=header)
    result.to_excel('combined_result.xlsx', index=False)
    print(f'combined_result.xlsx has been generated successfully，include {len(all_rows)} products.')


if __name__ == '__main__':
    main()
