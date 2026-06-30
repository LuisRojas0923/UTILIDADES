import zipfile
import re
import xml.etree.ElementTree as ET

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"

# Let's read sharedStrings
with zipfile.ZipFile(file_path, 'r') as z:
    ss_data = z.read("xl/sharedStrings.xml")
    root_ss = ET.fromstring(ss_data)
    # The namespace is usually http://schemas.openxmlformats.org/spreadsheetml/2006/main
    # Let's remove namespace for easier access
    ss_clean = re.sub(r' xmlns="[^"]+"', '', ss_data.decode("utf-8", errors="ignore"))
    root_ss = ET.fromstring(ss_clean)
    shared_strings = [t.text for t in root_ss.findall('.//t')]

    sheet_data = z.read("xl/worksheets/sheet18.xml").decode("utf-8", errors="ignore")
    
    # Print row 1 cells (headers)
    print("--- Row 1 (Headers) ---")
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
        cell_ref = f"{col}1"
        m = re.search(r'<c r="' + cell_ref + r'"([^>]*)>(.*?)</c>|<c r="' + cell_ref + r'"([^>]*)\/>', sheet_data, re.DOTALL)
        if m:
            inner = m.group(2) if m.group(2) else ""
            attrs = m.group(1) if m.group(1) else m.group(3)
            
            v_match = re.search(r'<v>(.*?)</v>', inner) if inner else None
            val = v_match.group(1) if v_match else "No value"
            
            # If string type, lookup in shared strings
            if 't="s"' in attrs and val != "No value":
                str_val = shared_strings[int(val)]
            else:
                str_val = val
                
            print(f"Cell {cell_ref}: {str_val}")
