import os
import zipfile
import re

original_file = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"
fixed_file = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026_FINAL_V2.xlsx"

def fix():
    last_detected_row = 1
    with zipfile.ZipFile(original_file, 'r') as zin:
        with zipfile.ZipFile(fixed_file, 'w', zipfile.ZIP_DEFLATED) as zout:
            items = zin.infolist()
            
            # 1. Process Sheet
            for item in items:
                if item.filename == 'xl/worksheets/sheet1.xml':
                    content = zin.read(item.filename)
                    row_pattern = re.compile(br'<row r="(\d+)"')
                    val_pattern = re.compile(br'<v>[^0<][^<]*</v>|<t>|<is>')
                    
                    row_matches = list(re.finditer(br'<row r="(\d+)"[^>]*>(.*?)</row>', content, re.DOTALL))
                    last_row = 1
                    for m in reversed(row_matches):
                        if val_pattern.search(m.group(2)):
                            last_row = int(m.group(1))
                            break
                    
                    last_detected_row = last_row
                    print(f"Detected last row: {last_row}")
                    
                    header_end = content.find(b'<sheetData>')
                    head = content[:header_end].decode('utf-8', errors='ignore')
                    head = re.sub(r'dimension ref="([A-Z0-9]+):[A-Z0-9]+"', rf'dimension ref="\1:CY{last_row}"', head)
                    
                    new_xml = head.encode('utf-8') + b'<sheetData>'
                    for m in row_matches:
                        if int(m.group(1)) <= last_row:
                            new_xml += m.group(0)
                        else: break
                    
                    footer_start = content.find(b'</sheetData>')
                    new_xml += content[footer_start:]
                    zout.writestr(item.filename, new_xml)
            
            # 2. Process Tables
            for item in items:
                if item.filename.startswith('xl/tables/'):
                    content = zin.read(item.filename).decode('utf-8', errors='ignore')
                    content = re.sub(r'ref="([A-Z]+[0-9]+):[A-Z]+[0-9]+"', rf'ref="\1:CU{last_detected_row}"', content)
                    content = re.sub(r'<autoFilter ref="([A-Z]+[0-9]+):[A-Z]+[0-9]+"', rf'<autoFilter ref="\1:CU{last_detected_row}"', content)
                    zout.writestr(item.filename, content.encode('utf-8'))
                elif item.filename == 'xl/calcChain.xml':
                    continue
                elif item.filename == 'xl/worksheets/sheet1.xml':
                    continue
                else:
                    zout.writestr(item.filename, zin.read(item.filename))

    print(f"Final file created: {fixed_file}")

if __name__ == "__main__":
    fix()
