import zipfile
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED.xlsm"

with zipfile.ZipFile(file_path, 'r') as z:
    sheet_data = z.read("xl/worksheets/sheet2.xml").decode("utf-8", errors="ignore")
    
    # Check column FV on row 11 (FV11)
    cell_match = re.search(r'<c r="FV11"([^>]*)>(.*?)</c>', sheet_data, re.DOTALL)
    if cell_match:
        inner = cell_match.group(2)
        f_match = re.search(r'<f[^>]*>(.*?)</f>', inner)
        formula = f_match.group(1) if f_match else "No formula"
        attrs = cell_match.group(1)
        print(f"Cell FV11 attrs: {attrs}")
        print(f"Cell FV11 Formula: {formula}")
    else:
        print("Cell FV11 not found in sheet")
