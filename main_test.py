import pandas as pd
import pandas as pd
import openpyxl
import glob
import os
import numpy as np
import excel_organization_func as ef

adapter_path = "../excel/adapter"
connector_path = "../excel/connector"
cab_assem_path = "../excel/cable assembly"
load_path = "../excel/load"

# adapter_param_names = ['Connector 1 Type', 'Connector 1 Impedance', 'Connector 1 Polarity',
#                    'Connector 2 Type', 'Connector 2 Impedance', 'Connector 2 Polarity',
#                    'Connector Mount Method', 'Adapter Body Style', 'Frequency', 'Insertion Loss (dB)',
#                    'VSWR /Return Loss', 'Center Contact', 'Outer Contact', 'Body', 'Dielectric',
#                    'Temperature Range', 'Compliant']
# adapter_param_match = [r'Connector\s*1\s*Type', r'Connector\s*1\s*Impedance', r'Connector\s*1\s*Polarity',
#                     r'Connector\s*2\s*Type', r'Connector\s*2\s*Impedance', r'Connector\s*2\s*Polarity',
#                     r'Connector\s*Mount\s*Method', r'Adapter\s*Body\s*Style', r'Frequency', r'Insertion\s*Loss\s*\(dB\)',
#                     r'(?:VSWR\s*/\s*Return\s*Loss)|(?:Return\s*Loss\s*/\s*VSWR)', r'Cent(?:er|re)\s*Contact', r'Outer\s*Contact', r'Body',
#                     r'Dielectric', r'Temperature\s*Range', r'Compliant']
adapter_param_dict = {'Connector 1 Type': r'Connector\s*1\s*Type', 'Connector 1 Impedance': r'Connector\s*1\s*Impedance',
                   'Connector 1 Polarity': r'Connector\s*1\s*Polarity','Connector 2 Type': r'Connector\s*2\s*Type',
                   'Connector 2 Impedance': r'Connector\s*2\s*Impedance', 'Connector 2 Polarity':  r'Connector\s*2\s*Polarity',
                   'Connector Mount Method': r'Connector\s*Mount\s*Method', 'Adapter Body Style': r'Adapter\s*Body\s*Style',
                   'Frequency': r'Frequency', 'Insertion Loss (dB)': r'Insertion\s*Loss\s*\(dB\)',
                   'VSWR /Return Loss': r'(?:VSWR\s*/\s*Return\s*Loss)|(?:Return\s*Loss\s*/\s*VSWR)',
                   'Center Contact': r'Cent(?:er|re)\s*Contact', 'Outer Contact': r'Outer\s*Contact',
                   'Body': r'Body', 'Dielectric': r'Dielectric','Temperature Range': r'Temperature\s*Range',
                   'Compliant': r'Compliant'}
# connector_param_names = ['Connector 1 Type', 'Connector 1 Impedance', 'Connector 1 Polarity', 'Body Style',
#                     'Connector Mount Method', 'Connector 2 Interface Type', 'Attachment Method','Frequency',
#                     'Insertion Loss (dB)','VSWR /Return Loss', 'Center Contact', 'Outer Contact', 'Body', 'Dielectric',
#                     'Temperature Range', 'Compliant']
# connector_param_match = [r'Connector\s*1?\s*Type', r'Connector\s*1?\s*Impedance', r'Connector\s*1?\s*Polarity', r'Body\s*Style',
#                     r'Connector\s*Mount\s*Method', r'Connector\s*2?\s*Interface\s*Type', r'Attachment\s*Method',r'Frequency',
#                     r'Insertion\s*Loss\s*\(dB\)',r'(?:VSWR\s*/\s*Return\s*Loss)|(?:Return\s*Loss\s*/\s*VSWR)', r'Cent(?:re|er)\s*Contact',
#                     r'Outer\s*Contact', r'Body', r'Dielectric', r'Temperature\s*Range', r'Compliant']
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
# cab_assem_param_names = ['Connector 1 Type','Connector 1 Body Style', 'Body Material and Plating','Connector 1 Mount Method',
#                          'Connector 2 Type','Connector 2 Body Style', 'Body Material and Plating','Connector 2 Mount Method',
#                          'Cable Type', 'Impedance', 'Frequency', 'Return Loss /VSWR', 'Phase Stability vs. Flexure',
#                          'Amplitude Stability','Shielding Effectiveness','Phase Matching','Signal Delay','Power Handling',
#                          'Temperature Range','Compliant']
# cab_assem_param_match = [r'Connector\s*1\s*Type',r'Connector\s*1\s*Body\s*Style', r'Body\s*Material\s*and\s*Plating',r'Connector\s*1\s*Mount\s*Method',
#                          r'Connector\s*2\s*Type',r'Connector\s*2\s*Body\s*Style', r'Body\s*Material\s*and\s*Plating',r'Connector\s*2\s*Mount\s*Method',
#                          r'Cable\s*Type', r'Impedance', r'Frequency', r'(?:VSWR\s*/\s*Return Loss)|(?:Return Loss\s*/\s*VSWR)',
#                          r'Phase\s*Stability\s*vs.\s*Flexure', r'Amplitude\s*Stability',r'Shielding\s*Effectiveness',r'Phase\s*Matching',
#                          r'Signal\s*Delay',r'Power\s*Handling', r'Temperature\s*Range',r'Compliant']
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
# load_param_names = ['Connector 1 Type', 'Connector 1 Impedance', 'Connector 1 Polarity','Frequency','Insertion Loss (dB)',
#                     'VSWR /Return Loss','Power','Center Contact','Outer Contact','Body','Dielectric','Temperature Range','Compliant']
# load_param_match = [r'Connector\s*1?\s*Type', r'Connector\s*1?\s*Impedance', r'Connector\s*1?\s*Polarity',r'Frequency',
#                     r'Insertion\s*Loss\s*\(dB\)',r'(?:VSWR\s*/\s*Return\s*Loss)|(?:Return\s*Loss\s*/\s*VSWR)', r'Power',
#                     r'Cent(?:re|er)\s*Contact',r'Outer\s*Contact',r'Body',r'Dielectric',r'Temperature\s*Range',r'Compliant']
load_param_dict = {'Connector 1 Type': r'Connector\s*1?\s*Type', 'Connector 1 Impedance': r'Connector\s*1?\s*Impedance',
                'Connector 1 Polarity': r'Connector\s*1?\s*Polarity','Frequency': r'Frequency',
                'Insertion Loss (dB)':r'Insertion\s*Loss\s*\(dB\)',
                'VSWR /Return Loss': r'(?:VSWR\s*/\s*Return\s*Loss)|(?:Return\s*Loss\s*/\s*VSWR)',
                'Power': r'Power','Center Contact': r'Cent(?:re|er)\s*Contact','Outer Contact': r'Outer\s*Contact',
                'Body': r'Body','Dielectric': r'Dielectric','Temperature Range': r'Temperature\s*Range',
                'Compliant': r'Compliant'}

adapter_result = ef.extract_from_folder(adapter_path, adapter_param_dict)
connector_result = ef.extract_from_folder(connector_path, connector_param_dict)
cab_assem_result = ef.extract_from_folder(cab_assem_path, cab_assem_param_dict)
load_result = ef.extract_from_folder(load_path, load_param_dict)

with pd.ExcelWriter('../excel/Combined_result.xlsx', engine='openpyxl') as writer:
    adapter_result.to_excel(writer, index=True, sheet_name='Adapter')
    connector_result.to_excel(writer, index=True, sheet_name='Connector')
    cab_assem_result.to_excel(writer, index=True, sheet_name='Cable Assembly')
    load_result.to_excel(writer, index=True, sheet_name='Load')
# print(f'combined_result.xlsx has been generated successfully，include {len(product_info_rows)} products.')