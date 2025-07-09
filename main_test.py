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

adapter_param_names = ['Connector 1 Type', 'Connector 1 Impedance', 'Connector 1 Polarity',
                   'Connector 2 Type', 'Connector 2 Impedance', 'Connector 2 Polarity',
                   'Connector Mount Method', 'Adapter Body Style', 'Frequency', 'Insertion Loss (dB)',
                   'VSWR /Return Loss', 'Center Contact', 'Outer Contact', 'Body', 'Dielectric',
                   'Temperature Range', 'Compliant']
adapter_param_match = [r'Connector\s*1\s*Type', r'Connector\s*1\s*Impedance', r'Connector\s*1\s*Polarity',
                    r'Connector\s*2\s*Type', r'Connector\s*2\s*Impedance', r'Connector\s*2\s*Polarity',
                    r'Connector Mount Method', r'Adapter Body Style', r'Frequency', r'Insertion Loss \(dB\)',
                    r'(?:VSWR\s*/\s*Return Loss)|(?:Return Loss\s*/\s*VSWR)', r'Cent(?:er|re) Contact', r'Outer Contact', r'Body',
                    r'Dielectric', r'Temperature Range', r'Compliant']
connector_param_names = ['Connector 1 Type', 'Connector 1 Impedance', 'Connector 1 Polarity', 'Body Style',
                    'Connector Mount Method', 'Connector 2 Interface Type', 'Attachment Method','Frequency',
                    'Insertion Loss (dB)','VSWR /Return Loss', 'Center Contact', 'Outer Contact', 'Body', 'Dielectric',
                    'Temperature Range', 'Compliant']
connector_param_match = [r'Connector\s*1\s*Type', r'Connector\s*1\s*Impedance', r'Connector\s*1\s*Polarity', r'Body Style',
                    r'Connector Mount Method', r'Connector\s*2\s*Interface Type', r'Attachment Method',r'Frequency',
                    r'Insertion Loss \(dB\)',r'(?:VSWR\s*/\s*Return Loss)|(?:Return Loss\s*/\s*VSWR)', r'Cent(?:re|er) Contact',
                    r'Outer Contact', r'Body', r'Dielectric', r'Temperature Range', r'Compliant']
cab_assem_param_names = ['Connector 1 Type','Connector 1 Body Style', 'Body Material and Plating','Connector 1 Mount Method',
                         'Connector 2 Type','Connector 2 Body Style', 'Body Material and Plating','Connector 2 Mount Method',
                         'Cable Type', 'Impedance', 'Frequency', 'Return Loss /VSWR', 'Phase Stability vs. Flexure',
                         'Amplitude Stability','Shielding Effectiveness','Phase Matching','Signal Delay','Power Handling',
                         'Temperature Range','Compliant']
cab_assem_param_match = [r'Connector\s*1\s*Type',r'Connector\s*1\s*Body\s*Style', r'Body\s*Material\s*and\s*Plating',r'Connector\s*1\s*Mount\s*Method',
                         r'Connector\s*2\s*Type',r'Connector\s*2\s*Body\s*Style', r'Body\s*Material\s*and\s*Plating',r'Connector\s*2\s*Mount\s*Method',
                         r'Cable\s*Type', r'Impedance', r'Frequency', r'(?:VSWR\s*/\s*Return Loss)|(?:Return Loss\s*/\s*VSWR)',
                         r'Phase\s*Stability\s*vs.\s*Flexure', r'Amplitude\s*Stability',r'Shielding\s*Effectiveness',r'Phase\s*Matching',
                         r'Signal\s*Delay',r'Power\s*Handling', r'Temperature\s*Range',r'Compliant']
load_param_names = ['Connector 1 Type', 'Connector 1 Impedance', 'Connector 1 Polarity','Frequency','Insertion Loss (dB)',
                    'VSWR /Return Loss','Power','Center Contact','Outer Contact','Body','Dielectric','Temperature Range','Compliant']
load_param_match = [r'Connector\s*1\s*Type', r'Connector\s*1\s*Impedance', r'Connector\s*1\s*Polarity',r'Frequency',
                    r'Insertion\s*Loss\s*\(dB\)',r'(?:VSWR\s*/\s*Return Loss)|(?:Return Loss\s*/\s*VSWR)', r'Power',
                    r'Cent(?:re|er) Contact',r'Outer Contact',r'Body',r'Dielectric',r'Temperature Range',r'Compliant']

adapter_result = ef.extract_from_folder(adapter_path, adapter_param_names, adapter_param_match)
connector_result = ef.extract_from_folder(connector_path, connector_param_names, connector_param_match)
cab_assem_result = ef.extract_from_folder(cab_assem_path, cab_assem_param_names, cab_assem_param_match)
load_result = ef.extract_from_folder(load_path, load_param_names, load_param_match)

with pd.ExcelWriter('../excel/Combined_result.xlsx', engine='openpyxl') as writer:
    adapter_result.to_excel(writer, index=True, sheet_name='Adapter')
    connector_result.to_excel(writer, index=True, sheet_name='Connector')
    cab_assem_result.to_excel(writer, index=True, sheet_name='Cable Assembly')
    load_result.to_excel(writer, index=True, sheet_name='Load')
# print(f'combined_result.xlsx has been generated successfully，include {len(product_info_rows)} products.')