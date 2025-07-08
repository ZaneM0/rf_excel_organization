import pandas as pd
import pandas as pd
import openpyxl
import glob
import os
import numpy as np
import excel_organization_func as ef

adapter_path = "../excel/adapter"
connector_path = "../excel/connector"

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

adapter_result = ef.extract_from_folder(adapter_path, adapter_param_names, adapter_param_match)
connector_result = ef.extract_from_folder(connector_path, connector_param_names, connector_param_match)
with pd.ExcelWriter('../excel/Combined_result.xlsx', engine='openpyxl') as writer:
    adapter_result.to_excel(writer, index=True, sheet_name='Adapter')
    connector_result.to_excel(writer, index=True, sheet_name='Connector')
# print(f'combined_result.xlsx has been generated successfully，include {len(product_info_rows)} products.')