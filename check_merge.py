import zipfile
import re

with zipfile.ZipFile('Anexo 1Q Abril 2026.xlsx', 'r') as zin:
    content = zin.read('xl/worksheets/sheet1.xml').decode('utf-8', errors='ignore')
    merge_cells = re.search(r'<mergeCells count="\d+">(.*?)</mergeCells>', content, re.DOTALL)
    if merge_cells:
        print(merge_cells.group(0)[:1000])
    else:
        print("No mergeCells found")
