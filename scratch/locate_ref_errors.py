import zipfile
import re
import xml.etree.ElementTree as ET

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"

def locate_ref_errors(sheet_file, sheet_name):
    print(f"\nLocating #REF! errors in {sheet_name} ({sheet_file}):")
    with zipfile.ZipFile(file_path, 'r') as z:
        sheet_data = z.read(sheet_file).decode("utf-8", errors="ignore")
        
        # We want to find cells `<c r="REF" ...><f>...#REF!...</f></c>`
        # Let's search using finditer
        cell_matches = re.finditer(r'<c r="([A-Z]+[0-9]+)"([^>]*)>(.*?)</c>', sheet_data, re.DOTALL)
        
        count = 0
        for m in cell_matches:
            ref = m.group(1)
            attrs = m.group(2)
            inner = m.group(3)
            
            if '<f' in inner and '#REF' in inner:
                count += 1
                f_content = re.search(r'<f[^>]*>(.*?)</f>', inner)
                formula = f_content.group(1) if f_content else "Unknown"
                # Decode XML entities
                formula = formula.replace('&quot;', '"').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
                if count <= 25: # print first 25
                    print(f"  - Cell {ref}: {formula}")
                
        print(f"Total #REF! formulas in {sheet_name}: {count}")

locate_ref_errors("xl/worksheets/sheet2.xml", "AUDITORIA NOMINA")
locate_ref_errors("xl/worksheets/sheet18.xml", "NOM_COMPLETA")
