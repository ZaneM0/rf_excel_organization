import pandas as pd
import re
import openpyxl
import glob
import os
import numpy as np
import excel_organization_func as ef

undocumented_path = '../excel/Combined_result.xlsx'
pim_header_path = '../excel/pim_header.csv'
adapter_df = pd.read_excel(undocumented_path, header=0, sheet_name='Adapter')
pim_param_name_ls = pd.read_csv(pim_header_path, header=0).columns.tolist()
adapter_param_name_ls = adapter_df.columns.tolist()

adapter_header_mapping_dict = {pim_param_name: '' for pim_param_name in pim_param_name_ls}
adapter_header_mapping_dict['Identifier'] = 'Product Name'
adapter_header_mapping_dict['Connector 1 Series'] = 'Connector 1 Type'
adapter_header_mapping_dict['Connector 2 Series'] = 'Connector 2 Type'
adapter_header_mapping_dict['Connector 1 Gender'] = 'Connector 1 Type'
adapter_header_mapping_dict['Connector 2 Gender'] = 'Connector 2 Type'
adapter_header_mapping_dict['Connector 1 Impedance (Ohm)'] = 'Connector 1 Impedance'
adapter_header_mapping_dict['Connector 2 Impedance (Ohm)'] = 'Connector 2 Impedance'
adapter_header_mapping_dict['Connector 1 Polarity'] = 'Connector 1 Polarity'
adapter_header_mapping_dict['Connector 2 Polarity'] = 'Connector 2 Polarity'
adapter_header_mapping_dict['Connector 1 Mount Method'] = 'Connector Mount Method'
adapter_header_mapping_dict['Connector 2 Mount Method'] = 'Connector Mount Method'
adapter_header_mapping_dict['Body Style'] = 'Adapter Body Style'
adapter_header_mapping_dict['Frequency'] = 'Frequency'
adapter_header_mapping_dict['Insertion Loss'] = 'Insertion Loss (dB)'
adapter_header_mapping_dict['VSWR / Return Loss'] = 'VSWR /Return Loss'
adapter_header_mapping_dict['Connector 1 Body Material'] = 'Body'
adapter_header_mapping_dict['Connector 2 Body Material'] = 'Body'
adapter_header_mapping_dict['Connector 1 Body Plating'] = 'Body'
adapter_header_mapping_dict['Connector 2 Body Plating'] = 'Body'
adapter_header_mapping_dict['Operating Temperature Range'] = 'Temperature Range'
adapter_header_mapping_dict['RoHS Compliant'] = 'Compliant'

write_to_pim_excel_ls =[]
for row_index in adapter_df.index:
    pim_product_info_ls = []
    for pim_param_name in pim_param_name_ls:
        if not adapter_header_mapping_dict[pim_param_name]:
            param_value = ''
        else:
            param_value = adapter_df.loc[row_index, adapter_header_mapping_dict[pim_param_name]]
        pim_product_info_ls.append(param_value)
    write_to_pim_excel_ls.append(pim_product_info_ls)

output_pim_df = pd.DataFrame(write_to_pim_excel_ls, columns=pim_param_name_ls)
# output_pim_df.set_index(pim_param_name_ls[0], inplace=True)
for col in ['Connector 1 Series', 'Connector 2 Series']:
    output_pim_df[col] = (
        output_pim_df[col]
        .astype(str)
        .str.strip()
        .str.replace(r'\s+(Male|Female)$', '', regex=True, flags=re.IGNORECASE)
        .replace('nan', np.nan)
    )

for col in ['Connector 1 Gender', 'Connector 2 Gender']:
    output_pim_df[col] = (
        output_pim_df[col]
        .astype(str)
        .str.strip()
        .str.extract(r'(Male|Female)', flags=re.IGNORECASE)[0]
        .replace('nan', np.nan)
    )

for col in ['Connector 1 Impedance (Ohm)', 'Connector 2 Impedance (Ohm)']:
    output_pim_df[col] = (
        output_pim_df[col]
        .astype(str)
        .str.strip()
        .str.replace(r'\s*(Ohms)$','', regex=True, flags=re.IGNORECASE)
        .replace('nan', np.nan)
    )

def adapter_body_material_info_process (element: str):
    stainless_steel_re = re.compile(r'Stainless\s*Steel', flags=re.IGNORECASE)
    brass_re = re.compile(r'Brass', flags=re.IGNORECASE)
    beryllium_copper_re = re.compile(r'Copper', flags=re.IGNORECASE)
    kovar_re = re.compile(r'Kovar', flags=re.IGNORECASE)
    cube_re = re.compile(r'CuBe', flags=re.IGNORECASE)
    if stainless_steel_re.search(element):
        return 'Stainless Steel'
    elif brass_re.search(element):
        return 'Brass'
    elif beryllium_copper_re.search(element):
        return 'Beryllium Copper'
    elif kovar_re.search(element):
        return 'Kovar'
    elif cube_re.search(element):
        return 'CuBe'
    else:
        return 'nan'

def adapter_body_plating_info_process (element: str):
    stainless_steel_re = re.compile(r'Stainless\s*Steel', flags=re.IGNORECASE)
    gold_re = re.compile(r'Gold', flags=re.IGNORECASE)
    nickel_re = re.compile(r'Nickel|Ni', flags=re.IGNORECASE)
    silver_re = re.compile(r'Silver', flags=re.IGNORECASE)
    tri_metal_re = re.compile(r'Tri\s*-?\s*Metal', flags=re.IGNORECASE)
    if stainless_steel_re.search(element):
        return 'Passivated'
    elif gold_re.search(element):
        return 'Gold'
    elif nickel_re.search(element):
        return 'Nickel'
    elif silver_re.search(element):
        return 'Silver'
    elif tri_metal_re.search(element):
        return 'Tri-Metal'
    else:
        return 'nan'


for col in ['Connector 1 Body Material', 'Connector 2 Body Material']:
    output_pim_df[col] = (
        output_pim_df[col]
        .astype(str)
        .apply(adapter_body_material_info_process)
        .replace('nan', np.nan)
    )

for col in ['Connector 1 Body Plating', 'Connector 2 Body Plating']:
    output_pim_df[col] = (
        output_pim_df[col]
        .astype(str)
        .apply(adapter_body_plating_info_process)
        .replace('nan', np.nan)
    )

    output_pim_df['RoHS Compliant'] = 'Yes'
    output_pim_df['Gwave PN'] = output_pim_df['Identifier']
    output_pim_df['Flexi PN'] = (
        output_pim_df['Identifier']
        .astype(str)
        .apply(lambda x: 'FR1-' + x)
    )

output_pim_df.to_excel('../excel/output_to_pim/adapter_pim_output_excel.xlsx', index=False)
print()
print('done!')
#KeyError
