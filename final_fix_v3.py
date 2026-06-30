import os
import zipfile
import re

original_file = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"
fixed_file = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026_FINAL_V3.xlsx"

def fix():
    last_row = 98 # We already know this from previous analysis
    with zipfile.ZipFile(original_file, 'r') as zin:
        with zipfile.ZipFile(fixed_file, 'w', zipfile.ZIP_DEFLATED) as zout:
            items = zin.infolist()
            
            for item in items:
                if item.filename == 'xl/worksheets/sheet1.xml':
                    print("Processing sheet1.xml...")
                    content = zin.read(item.filename)
                    
                    # Find where row 99 starts
                    # We look for <row r="99" or <row r="99">
                    target = b'<row r="99"'
                    pos = content.find(target)
                    
                    if pos != -1:
                        # Find where sheetData starts to fix dimension
                        sd_pos = content.find(b'<sheetData>')
                        head = content[:sd_pos].decode('utf-8', errors='ignore')
                        head = re.sub(r'dimension ref="([A-Z0-9]+):[A-Z0-9]+"', rf'dimension ref="\1:CY{last_row}"', head)
                        
                        # Find where sheetData ends to preserve footer
                        # We skip from 'pos' until </sheetData>
                        footer_pos = content.find(b'</sheetData>', pos)
                        
                        new_xml = head.encode('utf-8') + b'<sheetData>' + content[sd_pos+11:pos] + content[footer_pos:]
                        zout.writestr(item.filename, new_xml)
                    else:
                        zout.writestr(item.filename, content)
                
                elif item.filename.startswith('xl/tables/'):
                    content = zin.read(item.filename).decode('utf-8', errors='ignore')
                    content = re.sub(r'ref="([A-Z]+[0-9]+):[A-Z]+[0-9]+"', rf'ref="\1:CU{last_row}"', content)
                    content = re.sub(r'<autoFilter ref="([A-Z]+[0-9]+):[A-Z]+[0-9]+"', rf'<autoFilter ref="\1:CU{last_row}"', content)
                    zout.writestr(item.filename, content.encode('utf-8'))
                elif item.filename == 'xl/calcChain.xml':
                    continue
                else:
                    zout.writestr(item.filename, zin.read(item.filename))

    print(f"Final file created: {fixed_file}")

if __name__ == "__main__":
    fix()
