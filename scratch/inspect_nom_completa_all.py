import zipfile
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"

with zipfile.ZipFile(file_path, 'r') as z:
    sheet_data = z.read("xl/worksheets/sheet18.xml").decode("utf-8", errors="ignore")
    
    # Find all cells in column N (matching N followed by numbers)
    matches = re.finditer(r'<c r="(N[0-9]+)"([^>]*)>(.*?)</c>|<c r="(N[0-9]+)"([^>]*)\/>', sheet_data, re.DOTALL)
    
    print("Found cells in column N:")
    for m in matches:
        ref = m.group(1) if m.group(1) else m.group(4)
        inner = m.group(3) if m.group(1) else ""
        
        f_match = re.search(r'<f[^>]*>(.*?)</f>', inner) if inner else None
        formula = f_match.group(1) if f_match else "No formula"
        
        v_match = re.search(r'<v>(.*?)</v>', inner) if inner else None
        val = v_match.group(1) if v_match else "No value"
        
        print(f"Cell {ref}: Value={val}, Formula={formula[:120]}")
