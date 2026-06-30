import zipfile
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"

with zipfile.ZipFile(file_path, 'r') as z:
    sheet_data = z.read("xl/worksheets/sheet18.xml").decode("utf-8", errors="ignore")
    
    # Let's inspect row 2 (headers) and row 311 for columns M and N, and also row 3 to see what is there
    for row in [2, 3, 311]:
        print(f"\n--- Row {row} ---")
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
            cell_ref = f"{col}{row}"
            m = re.search(r'<c r="' + cell_ref + r'"([^>]*)>(.*?)</c>|<c r="' + cell_ref + r'"([^>]*)\/>', sheet_data, re.DOTALL)
            if m:
                inner = m.group(2) if m.group(2) else ""
                attrs = m.group(1) if m.group(1) else m.group(3)
                
                f_match = re.search(r'<f[^>]*>(.*?)</f>', inner) if inner else None
                formula = f_match.group(1) if f_match else "No formula"
                
                v_match = re.search(r'<v>(.*?)</v>', inner) if inner else None
                val = v_match.group(1) if v_match else "No value"
                
                print(f"Cell {cell_ref}: Val={val}, Formula={formula[:60]}")
