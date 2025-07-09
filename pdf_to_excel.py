import pandas as pd
import glob
import os
import re
from PyPDF2 import PdfReader

def extract_data_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ''
    for page in reader.pages:
        text += page.extract_text() or ''
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    # 1. 商品类型：使用模糊匹配关键词
    product_type = 'N/A'
    patterns = {
        'connector': re.compile(r'\bconnector\b', re.IGNORECASE),
        'adapter': re.compile(r'\badapt(?:e|o)r\b', re.IGNORECASE),       # adapter 或 adaptor
        'cable assembly': re.compile(r'\bcable\s*assembly\b', re.IGNORECASE),  # 任意空格
        'load': re.compile(r'\bload\b', re.IGNORECASE),
    }
    header_snippet = ' '.join(lines[:10])
    for name, pat in patterns.items():
        if pat.search(header_snippet):
            product_type = name
            break

    # 2. 产品编号：在“datasheet”行之后的几行中定位
    product_number = 'N/A'
    idx_ds = next((i for i, ln in enumerate(lines) if 'datasheet' in ln.lower()), None)
    search_start = idx_ds + 1 if idx_ds is not None else 0
    for ln in lines[search_start:search_start+5]:
        if re.match(r'^[A-Z0-9][A-Z0-9._/\-]{1,}[A-Z0-9]$', ln):
            product_number = ln
            break

    # 3. 参数：自动抓取所有含有至少两个空格分隔的行
    params = {}
    collect_start = idx_ds + 1 if idx_ds is not None else 0
    for ln in lines[collect_start:]:
        parts = re.split(r'\s{2,}', ln)
        if len(parts) == 2:
            name, value = parts
            params[name] = value

    return product_type, product_number, params

def main():
    for pdf_file in glob.glob("*.pdf"):
        ptype, pnum, params = extract_data_from_pdf(pdf_file)
        rows = []
        # 第一行
        rows.append([f"{ptype} datasheet"])
        # 第二行
        rows.append(["product name", pnum])
        # 第三行空
        rows.append([])
        # 从第四行开始：参数名、空列、参数值
        for name, value in params.items():
            rows.append([name, "", value])

        df = pd.DataFrame(rows)
        out_file = os.path.splitext(pdf_file)[0] + ".xlsx"
        df.to_excel(out_file, index=False, header=False)
        print(f"Generated: {out_file}")

if __name__ == "__main__":
    main()