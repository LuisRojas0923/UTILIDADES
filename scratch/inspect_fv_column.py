import zipfile
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED.xlsm"

with zipfile.ZipFile(file_path, 'r') as z:
    sheet_data = z.read("xl/worksheets/sheet2.xml").decode("utf-8", errors="ignore")
    
    # Check column FV on row 195
    row_block_match = re.search(r'<row r="195"[^>]*>(.*?)</row>', sheet_data, re.DOTALL)
    if row_block_match:
        row_content = row_block_match.group(1)
        # Search for cell ref FV195
        cell_match = re.search(r'<c r="FV195"([^>]*)>(.*?)</c>', row_content, re.DOTALL)
        if cell_match:
            inner = cell_match.group(2)
            f_match = re.search(r'<f[^>]*>(.*?)</f>', inner)
            formula = f_match.group(1) if f_match else "No formula"
            print(f"Cell FV195: Formula={formula}")
        else:
            print("Cell FV195 not found in row 195")
    else:
        print("Row 195 not found")
