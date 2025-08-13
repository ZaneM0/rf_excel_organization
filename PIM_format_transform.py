import pandas as pd
import re
import openpyxl
import glob
import os
import numpy as np
import excel_organization_func as ef

excel_to_be_transformed_path = '../excel/Combined_result.xlsx'
adapter_pim_header_path = '../excel/pim_header/adapter_pim_header.csv'
connector_pim_header_path = '../excel/pim_header/connector_pim_header.csv'
cable_assembly_pim_header_path = '../excel/pim_header/cable_assembly_pim_header.csv'

adapter_df = pd.read_excel(excel_to_be_transformed_path, header=0, sheet_name='Adapter')
connector_df = pd.read_excel(excel_to_be_transformed_path, header=0, sheet_name='Connector')
cable_assembly_df = pd.read_excel(excel_to_be_transformed_path, header=0, sheet_name='Cable Assembly')
adapter_header_ls = pd.read_csv(adapter_pim_header_path, header=0).columns.tolist()
connector_header_ls = pd.read_csv(connector_pim_header_path, header=0).columns.tolist()
cable_assembly_header_ls = pd.read_csv(connector_pim_header_path, header=0).columns.tolist()
#adapter_header_param_counter = {param_name: 0 for param_name in adapter_header}
# adapter_header_ls = []
# for param_name in adapter_header:
#     if not adapter_header_param_counter[param_name]:
#         adapter_header_ls.append(param_name)
#         adapter_header_param_counter[param_name] += 1
# print(adapter_header)
# print(adapter_header_ls)
#adapter_param_name_ls = adapter_df.columns.tolist()


def connector_series_info_process (element: str):
    series_patterns = [
        (re.compile(r'1\.0/2\.3', re.IGNORECASE), '1.0/2.3'),
        (re.compile(r'0\.8\s*mm$', re.IGNORECASE), '0.8mm'),
        (re.compile(r'1(\.0)?\s*mm\b', re.IGNORECASE), '1.0mm'),
        (re.compile(r'1\.85\s*mm$', re.IGNORECASE), '1.85mm'),
        (re.compile(r'1\.85\s*mm\s*NMD\b', re.IGNORECASE), '1.85mm NMD'),
        (re.compile(r'2\.4\s*mm$', re.IGNORECASE), '2.4mm'),
        (re.compile(r'2\.4\s*mm\s*NMD\b', re.IGNORECASE), '2.4mm NMD'),
        (re.compile(r'2\.92\s*mm$', re.IGNORECASE), '2.92mm'),
        (re.compile(r'2\.92\s*mm\s*NMD\b', re.IGNORECASE), '2.92 NMD'),
        (re.compile(r'3\.5\s*mm$', re.IGNORECASE), '3.5mm'),
        (re.compile(r'3\.5\s*mm\s*NMD\b', re.IGNORECASE), '3.5mm NMD'),
        (re.compile(r'4\.3\s*-\s*10', re.IGNORECASE), '4.3-10'),
        (re.compile(r'7\s*/\s*16\s*DIN\b', re.IGNORECASE), '7/16 DIN'),
        (re.compile(r'7\s*mm\b', re.IGNORECASE), '7mm'),
        (re.compile(r'\bBMA\b', re.IGNORECASE), 'BMA'),
        (re.compile(r'\bBNC\b', re.IGNORECASE), 'BNC'),
        (re.compile(r'\bG3PO\b|\bSMPS\b', re.IGNORECASE), 'G3PO(SMPS)'),
        (re.compile(r'\bGPPO\b|\bMini\s*-\s*SMP\b', re.IGNORECASE), 'GPPO(Mini-SMP)'),
        (re.compile(r'\bGPO\b|\b(?<!Mini-)SMP\b', re.IGNORECASE), 'GPO(SMP)'),
        (re.compile(r'\bMCX\b', re.IGNORECASE), 'MCX'),
        (re.compile(r'\bMMCX\b', re.IGNORECASE), 'MMCX'),
        (re.compile(r'\bQuick\s*N\b', re.IGNORECASE), 'Quick N'),
        (re.compile(r'\bN\b', re.IGNORECASE), 'N'),
        (re.compile(r'\bQuick\s*SMA\b', re.IGNORECASE), 'Quick SMA'),
        (re.compile(r'\bSMA\b', re.IGNORECASE), 'SMA'),
        (re.compile(r'\bSMB\b', re.IGNORECASE), 'SMB'),
        (re.compile(r'\bSMC\b', re.IGNORECASE), 'SMC'),
        (re.compile(r'\bQuick\s*SSMA\b', re.IGNORECASE), 'Quick SSMA'),
        (re.compile(r'\bSSMA\b', re.IGNORECASE), 'SSMA'),
        (re.compile(r'\bSSMB\b', re.IGNORECASE), 'SSMB'),
        (re.compile(r'\bSSMC\b', re.IGNORECASE), 'SSMC'),
        (re.compile(r'\bTNC\b', re.IGNORECASE), 'TNC'),
        (re.compile(r'\bQMA\b', re.IGNORECASE), 'QMA'),
        (re.compile(r'\bUHF\b', re.IGNORECASE), 'UHF'),
        (re.compile(r'\bMMPX\b', re.IGNORECASE), 'MMPX'),
        (re.compile(r'\bIPX1\b', re.IGNORECASE), 'IPX1'),
        (re.compile(r'\bIPX4\b', re.IGNORECASE), 'IPX4')
    ]
    for series_pattern, series_name in series_patterns:
        if series_pattern.search (element):
            return series_name

    return element

def body_material_info_process (element: str):
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
        return element

def body_plating_info_process (element: str):
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
        return element

def adapter_pim_transform():
    adapter_header_mapping_dict = {pim_param_name: [] for pim_param_name in adapter_header_ls}
    #dicionary used to map the headers of original Excel file to pim headers
    # xxx_header_mapping_dict['original header name'] = ['corresponding PIM header name1', 'name2'(if exist), ...]
    adapter_header_mapping_dict['Identifier'] = ['Product Name']
    adapter_header_mapping_dict['Connector 1 Series'] = ['Connector 1 Type']
    adapter_header_mapping_dict['Connector 2 Series'] = ['Connector 2 Type']
    adapter_header_mapping_dict['Connector 1 Gender'] = ['Connector 1 Type']
    adapter_header_mapping_dict['Connector 2 Gender'] = ['Connector 2 Type']
    adapter_header_mapping_dict['Connector 1 Impedance (Ohm)'] = ['Connector 1 Impedance']
    adapter_header_mapping_dict['Connector 2 Impedance (Ohm)'] = ['Connector 2 Impedance']
    adapter_header_mapping_dict['Connector 1 Polarity'] = ['Connector 1 Polarity']
    adapter_header_mapping_dict['Connector 2 Polarity'] = ['Connector 2 Polarity']
    adapter_header_mapping_dict['Connector 1 Mount Method'] = ['Connector Mount Method']
    adapter_header_mapping_dict['Connector 2 Mount Method'] = ['Connector Mount Method']
    adapter_header_mapping_dict['Body Style'] = ['Adapter Body Style']
    adapter_header_mapping_dict['Frequency'] = ['Frequency']
    adapter_header_mapping_dict['Insertion Loss'] = ['Insertion Loss (dB)']
    adapter_header_mapping_dict['VSWR / Return Loss'] = ['VSWR /Return Loss']
    adapter_header_mapping_dict['Connector 1 Body Material'] = ['Body', 'Outer Contact']
    adapter_header_mapping_dict['Connector 2 Body Material'] = ['Body', 'Outer Contact']
    adapter_header_mapping_dict['Connector 1 Body Plating'] = ['Body', 'Outer Contact']
    adapter_header_mapping_dict['Connector 2 Body Plating'] = ['Body', 'Outer Contact']
    adapter_header_mapping_dict['Operating Temperature Range'] = ['Temperature Range']
    adapter_header_mapping_dict['RoHS Compliant'] = ['Compliant']

    write_to_pim_excel_ls =[]
    for row_index in adapter_df.index:
        pim_product_info_ls = []
        for pim_param_name in adapter_header_ls:
            param_value = ''
            if adapter_header_mapping_dict[pim_param_name]:
                for param_name in adapter_header_mapping_dict[pim_param_name]:
                    param_value = adapter_df.loc[row_index, param_name]
                    if param_value != '' and not pd.isna(param_value):
                        break
            pim_product_info_ls.append(param_value)
        write_to_pim_excel_ls.append(pim_product_info_ls)

    output_pim_df = pd.DataFrame(write_to_pim_excel_ls, columns=adapter_header_ls)
    # output_pim_df.set_index(adapter_header[0], inplace=True)
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


    for col in ['Connector 1 Body Material', 'Connector 2 Body Material']:
        output_pim_df[col] = (
            output_pim_df[col]
            .astype(str)
            .apply(body_material_info_process)
            .replace('nan', np.nan)
        )

    for col in ['Connector 1 Body Plating', 'Connector 2 Body Plating']:
        output_pim_df[col] = (
            output_pim_df[col]
            .astype(str)
            .apply(body_plating_info_process)
            .replace('nan', np.nan)
        )

    for col in ['Connector 1 Series', 'Connector 2 Series']:
        output_pim_df[col] = (
            output_pim_df[col]
            .astype(str)
            .apply(connector_series_info_process)
            .replace('nan', np.nan)
        )

    output_pim_df['RoHS Compliant'] = 'Yes'

    output_pim_df['Gwave PN'] = output_pim_df['Identifier']

    output_pim_df['Product Type'] = 'Adapter'

    output_pim_df['Type'] = 'Simple'

    output_pim_df['Vendor'] = 'Flexi RF Inc'

    output_pim_df['Tags'] = (
        output_pim_df.fillna('')
        .astype(str)
        .apply(lambda row: row['Connector 1 Series'] + ', ' + row['Connector 2 Series']
               if row['Connector 1 Series'] != row['Connector 2 Series'] else row['Connector 1 Series'],axis=1)
    )

    output_pim_df['Flexi PN'] = (
        output_pim_df['Identifier']
        .astype(str)
        .apply(lambda x: 'FR1-' + x)
    )

    output_pim_df.to_excel('../excel/output_to_pim/adapter_pim_output_excel.xlsx', index=False)

def main():
    adapter_pim_transform()
    print('done!')

if __name__ == '__main__':
    main()
#KeyError
