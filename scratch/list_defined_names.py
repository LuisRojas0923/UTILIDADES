import zipfile
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"

with zipfile.ZipFile(file_path, 'r') as z:
    workbook_xml = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
    
    # We want to extract <definedName name="...">refers_to</definedName>
    # Let's search using finditer
    dn_matches = re.finditer(r'<definedName\s+name="([^"]+)"([^>]*)>(.*?)</definedName>', workbook_xml, re.DOTALL)
    
    print("Found Defined Names in workbook:")
    for m in dn_matches:
        name = m.group(1)
        attrs = m.group(2)
        refers_to = m.group(3)
        # Check if localSheetId is present
        ls_match = re.search(r'localSheetId="(\d+)"', attrs)
        local_sheet = ls_match.group(1) if ls_match else "Global"
        
        # Clean refers_to
        refers_to = refers_to.replace('&quot;', '"').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        print(f"  - Name: '{name}' ({local_sheet}) -> {refers_to}")
