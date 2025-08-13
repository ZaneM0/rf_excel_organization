import pandas as pd
import openpyxl
import re
import glob
import os
import numpy as np
import excel_organization_func as ef

adapter_path = "../excel/adapter"
connector_path = "../excel/connector"
cab_assem_path = "../excel/cable assembly"
load_path = "../excel/load"
already_documented_products_path = "../excel/already_documented_products.csv"

documented_set = set()
documented_products_df = pd.read_csv(already_documented_products_path,header=0)
identifier_col = documented_products_df['Identifier']
for identifier in identifier_col:
    if identifier not in documented_set:
        documented_set.add(identifier)

adapter_param_dict = {'Connector 1 Type': r'Type', 'Connector 1 Impedance': r'Connector\s*1\s*Impedance',
                   'Connector 1 Polarity': r'Connector\s*1\s*Polarity','Connector 2 Type': r'Type',
                   'Connector 2 Impedance': r'Connector\s*2\s*Impedance', 'Connector 2 Polarity':  r'Connector\s*2\s*Polarity',
                   'Connector Mount Method': r'Connector\s*Mount\s*Method', 'Adapter Body Style': r'Adapter\s*Body\s*Style',
                   'Frequency': r'Frequency', 'Insertion Loss (dB)': r'Insertion\s*Loss\s*\(dB\)',
                   'VSWR /Return Loss': r'(?:VSWR\s*/\s*Return\s*Loss)|(?:Return\s*Loss\s*/\s*VSWR)|(?:Return\s*Loss)|(?:VSWR)',
                   'Center Contact': r'Cent(?:er|re)\s*Contact', 'Outer Contact': r'Outer\s*Contact',
                   'Body': r'^Body', 'Dielectric': r'Dielectric','Temperature Range': r'Temperature\s*Range',
                   'Compliant': r'Compliant'}

connector_param_dict = {'Connector 1 Type': r'Connector\s*1?\s*Type', 'Connector 1 Impedance': r'Connector\s*1?\s*Impedance',
                     'Connector 1 Polarity': r'Connector\s*1?\s*Polarity', 'Body Style': r'Body\s*Style',
                     'Connector Mount Method': r'Connector\s*Mount\s*Method',
                     'Connector 2 Interface Type': r'Connector\s*2?\s*Interface\s*Type',
                     'Attachment Method': r'Attachment\s*Method', 'Frequency': r'Frequency',
                     'Insertion Loss (dB)': r'Insertion\s*Loss\s*\(dB\)',
                     'VSWR /Return Loss': r'(?:VSWR\s*/\s*Return\s*Loss)|(?:Return\s*Loss\s*/\s*VSWR)',
                     'Center Contact': r'Cent(?:re|er)\s*Contact', 'Outer Contact': r'Outer\s*Contact',
                     'Body': r'Body', 'Dielectric': r'Dielectric', 'Temperature Range': r'Temperature\s*Range',
                     'Compliant': r'Compliant'}

cab_assem_param_dict = {'Connector 1 Type': r'Connector\s*1\s*Type','Connector 1 Body Style': r'Connector\s*1\s*Body\s*Style',
                    'Connector 1 Body Material and Plating': r'Body\s*Material\s*and\s*Plating',
                    'Connector 1 Mount Method': r'Connector\s*1\s*Mount\s*Method',
                    'Connector 2 Type': r'Connector\s*2\s*Type','Connector 2 Body Style': r'Connector\s*2\s*Body\s*Style',
                    'Connector 2 Body Material and Plating': r'Body\s*Material\s*and\s*Plating',
                    'Connector 2 Mount Method': r'Connector\s*2\s*Mount\s*Method',
                    'Cable Type': r'Cable\s*Type', 'Impedance': r'Impedance', 'Frequency': r'Frequency',
                    'Return Loss /VSWR': r'(?:VSWR\s*/\s*Return Loss)|(?:Return Loss\s*/\s*VSWR)',
                    'Phase Stability vs. Flexure': r'Phase\s*Stability\s*vs.\s*Flexure',
                    'Amplitude Stability': r'Amplitude\s*Stability','Shielding Effectiveness': r'Shielding\s*Effectiveness',
                    'Phase Matching': r'Phase\s*Matching','Signal Delay': r'Signal\s*Delay','Power Handling': r'Power\s*Handling',
                    'Temperature Range': r'Temperature\s*Range','Compliant': r'Compliant'}

load_param_dict = {'Connector 1 Type': r'Connector\s*1?\s*Type', 'Connector 1 Impedance': r'Connector\s*1?\s*Impedance',
                'Connector 1 Polarity': r'Connector\s*1?\s*Polarity','Frequency': r'Frequency',
                'Insertion Loss (dB)':r'Insertion\s*Loss\s*\(dB\)',
                'VSWR /Return Loss': r'(?:VSWR\s*/\s*Return\s*Loss)|(?:Return\s*Loss\s*/\s*VSWR)',
                'Power': r'Power','Center Contact': r'Cent(?:re|er)\s*Contact','Outer Contact': r'Outer\s*Contact',
                'Body': r'Body','Dielectric': r'Dielectric','Temperature Range': r'Temperature\s*Range',
                'Compliant': r'Compliant'}
def main():
    adapter_df = ef.extract_from_folder(adapter_path, adapter_param_dict,documented_set)
    adapter_result = ef.replace_first_char_if_not_digit(adapter_df,['Insertion Loss (dB)','VSWR /Return Loss'])

    connector_df = ef.extract_from_folder(connector_path, connector_param_dict,documented_set)
    connector_result = ef.replace_first_char_if_not_digit(connector_df, ['Insertion Loss (dB)','VSWR /Return Loss'])

    cab_assem_result = ef.extract_from_folder(cab_assem_path, cab_assem_param_dict, documented_set)
    cab_assem_result['Impedance'] = (
        cab_assem_result['Impedance']
        .fillna('')
        .astype(str)
        .str.strip()
        .apply(lambda x: re.search(r'\d{2}', x).group(0) if x != 'N/A' and re.search(r'\d{2}', x) else 'N/A')
    )

    load_result = ef.extract_from_folder(load_path, load_param_dict, documented_set)

    with pd.ExcelWriter('../excel/Combined_result.xlsx', engine='openpyxl') as writer:
        adapter_result.to_excel(writer, index=True, sheet_name='Adapter')
        connector_result.to_excel(writer, index=True, sheet_name='Connector')
        cab_assem_result.to_excel(writer, index=True, sheet_name='Cable Assembly')
        load_result.to_excel(writer, index=True, sheet_name='Load')
    print(f'combined_result.xlsx has been generated successfully!')

if __name__ == '__main__':
    main()



