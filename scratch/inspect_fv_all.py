import zipfile
import re
import xml.etree.ElementTree as ET

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED.xlsm"

with zipfile.ZipFile(file_path, 'r') as z:
    ss_data = z.read("xl/sharedStrings.xml")
    ss_clean = re.sub(r' xmlns="[^"]+"', '', ss_data.decode("utf-8", errors="ignore"))
    root_ss = ET.fromstring(ss_clean)
    shared_strings = [t.text for t in root_ss.findall('.//t')]

    sheet_data = z.read("xl/worksheets/sheet2.xml").decode("utf-8", errors="ignore")
    
    # Find all cells in column FV: FV followed by numbers
    matches = re.finditer(r'<c r="(FV[0-9]+)"([^>]*)>(.*?)</c>|<c r="(FV[0-9]+)"([^>]*)\/>', sheet_data, re.DOTALL)
    
    print("Cells in Column FV:")
    for m in matches:
        ref = m.group(1) if m.group(1) else m.group(4)
        inner = m.group(3) if m.group(1) else ""
        attrs = m.group(2) if m.group(1) else m.group(5)
        
        v_match = re.search(r'<v>(.*?)</v>', inner) if inner else None
        val = v_match.group(1) if v_match else "No value"
        
        f_match = re.search(r'<f[^>]*>(.*?)</f>', inner) if inner else None
        formula = f_match.group(1) if f_match else "No formula"
        
        if 't="s"' in attrs and val != "No value":
            str_val = shared_strings[int(val)]
        else:
            str_val = val
            
        print(f"Cell {ref}: Value={str_val}, Formula={formula[:120]}")
