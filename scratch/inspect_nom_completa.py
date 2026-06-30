import zipfile
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"

with zipfile.ZipFile(file_path, 'r') as z:
    sheet_data = z.read("xl/worksheets/sheet18.xml").decode("utf-8", errors="ignore")
    
    # Let's find cells in column N: N309, N310, N311, N312, N313
    cells = ['N309', 'N310', 'N311', 'N312', 'N313']
    for cell in cells:
        m = re.search(r'<c r="' + cell + r'"([^>]*)>(.*?)</c>', sheet_data, re.DOTALL)
        if m:
            inner = m.group(2)
            f_match = re.search(r'<f[^>]*>(.*?)</f>', inner)
            formula = f_match.group(1) if f_match else "No formula"
            v_match = re.search(r'<v>(.*?)</v>', inner)
            val = v_match.group(1) if v_match else "No value"
            print(f"Cell {cell}: Value={val}, Formula={formula}")
        else:
            print(f"Cell {cell}: Not found")
