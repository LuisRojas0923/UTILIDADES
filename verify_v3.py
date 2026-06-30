import zipfile
import re

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026_FINAL_V3.xlsx"

with zipfile.ZipFile(file_path, 'r') as zin:
    table = zin.read('xl/tables/table1.xml').decode('utf-8')
    print(f"Table Ref: {re.search(r'ref=\"([^\"]+)\"', table).group(1)}")
    sheet = zin.read('xl/worksheets/sheet1.xml').decode('utf-8', errors='ignore')
    print(f"Sheet Dimension: {re.search(r'dimension ref=\"([^\"]+)\"', sheet).group(1)}")
