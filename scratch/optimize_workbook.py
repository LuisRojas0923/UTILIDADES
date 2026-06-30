import zipfile
import os
import re

src_file = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"
dst_file = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED.xlsm"

def optimize():
    if not os.path.exists(src_file):
        print(f"Error: Source file not found: {src_file}")
        return False
        
    print(f"Opening source: {src_file}")
    print(f"Writing to: {dst_file}")
    
    replacements_sheet2 = 0
    removed_cell_sheet18 = False
    trimmed_rows_sheet9 = 0
    dimension_updated_sheet9 = False
    
    with zipfile.ZipFile(src_file, 'r') as zin:
        with zipfile.ZipFile(dst_file, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                filename = item.filename
                data = zin.read(filename)
                
                if filename == "xl/worksheets/sheet2.xml":
                    # AUDITORIA NOMINA
                    content = data.decode("utf-8", errors="ignore")
                    
                    # Target: #REF!&amp;PLANILLA_REGIONAL_1Q[NOVEDAD]
                    # We want to replace it with PLANILLA_REGIONAL_1Q[CEDULA]&amp;PLANILLA_REGIONAL_1Q[NOVEDAD]
                    target = "#REF!&amp;PLANILLA_REGIONAL_1Q[NOVEDAD]"
                    replacement = "PLANILLA_REGIONAL_1Q[CEDULA]&amp;PLANILLA_REGIONAL_1Q[NOVEDAD]"
                    
                    matches = len(re.findall(re.escape(target), content))
                    if matches > 0:
                        content = content.replace(target, replacement)
                        replacements_sheet2 = matches
                        print(f"Sheet 2 (AUDITORIA NOMINA): Replaced {matches} occurrences of broken reference.")
                    else:
                        print("Sheet 2: Target broken reference not found.")
                        
                    data = content.encode("utf-8")
                    
                elif filename == "xl/worksheets/sheet18.xml":
                    # NOM_COMPLETA
                    content = data.decode("utf-8", errors="ignore")
                    
                    # Target: remove cell N311
                    # A cell tag: <c r="N311" ...>...</c> or <c r="N311" .../>
                    cell_pattern = re.compile(r'<c r="N311"[^>]*>.*?</c>|<c r="N311"[^>]*\/>', re.DOTALL)
                    if cell_pattern.search(content):
                        content = cell_pattern.sub('', content)
                        removed_cell_sheet18 = True
                        print("Sheet 18 (NOM_COMPLETA): Removed cell N311.")
                    else:
                        print("Sheet 18: Cell N311 not found.")
                        
                    data = content.encode("utf-8")
                    
                elif filename == "xl/worksheets/sheet9.xml":
                    # PRELIQUIDADO SIIGO
                    content = data.decode("utf-8", errors="ignore")
                    
                    # 1. Update dimension from A7:X1525 to A7:X1342
                    # The tag looks like <dimension ref="A7:X1525"/>
                    dim_pattern = re.compile(r'<dimension ref="([^"]+):X1525"\s*/>|<dimension ref="A7:X1525"\s*/>')
                    if dim_pattern.search(content):
                        content = dim_pattern.sub(r'<dimension ref="A7:X1342"/>', content)
                        dimension_updated_sheet9 = True
                        print("Sheet 9 (PRELIQUIDADO SIIGO): Updated dimension to A7:X1342.")
                    else:
                        # try general replacement if it starts differently
                        content, count = re.subn(r'ref="A7:X1525"', 'ref="A7:X1342"', content)
                        if count > 0:
                            dimension_updated_sheet9 = True
                            print("Sheet 9: Updated dimension ref using fallback.")
                        else:
                            print("Sheet 9: Dimension tag ref='A7:X1525' not found.")
                    
                    # 2. Remove rows > 1342
                    # Find all row blocks: <row r="Y" ...> ... </row>
                    # We can use a replacer function in sub
                    row_pattern = re.compile(r'<row r="(\d+)"[^>]*>.*?</row>|<row r="(\d+)"[^>]*\/>', re.DOTALL)
                    
                    def row_replacer(match):
                        nonlocal trimmed_rows_sheet9
                        r_num_str = match.group(1) if match.group(1) else match.group(2)
                        r_num = int(r_num_str)
                        if r_num > 1342:
                            trimmed_rows_sheet9 += 1
                            return "" # remove this row block
                        return match.group(0) # keep it
                        
                    content = row_pattern.sub(row_replacer, content)
                    print(f"Sheet 9: Trimmed {trimmed_rows_sheet9} row blocks (rows > 1342).")
                    
                    data = content.encode("utf-8")
                
                zout.writestr(filename, data)
                
    print("\nOptimization completed successfully!")
    print(f"Summary of changes:")
    print(f"  - Sheet 2 (AUDITORIA NOMINA) formulas corrected: {replacements_sheet2}")
    print(f"  - Sheet 18 (NOM_COMPLETA) cell N311 removed: {removed_cell_sheet18}")
    print(f"  - Sheet 9 (PRELIQUIDADO SIIGO) rows removed: {trimmed_rows_sheet9}")
    print(f"  - Sheet 9 (PRELIQUIDADO SIIGO) dimension updated: {dimension_updated_sheet9}")
    return True

if __name__ == "__main__":
    optimize()
