import zipfile
import xml.etree.ElementTree as ET

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"

try:
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        with zip_ref.open('xl/sharedStrings.xml') as strings:
            # We don't want to parse the whole thing if it's huge, but it's only 11KB
            tree = ET.parse(strings)
            root = tree.getroot()
            # Index 189
            si_elements = root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')
            if len(si_elements) > 189:
                # Get the text inside t
                t = si_elements[189].find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                print(f"Shared String 189: '{t.text if t is not None else ''}'")
            else:
                print(f"Only {len(si_elements)} shared strings found.")
                
except Exception as e:
    print(f"Error: {e}")
