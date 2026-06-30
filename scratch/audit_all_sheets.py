import zipfile
import os
import xml.etree.ElementTree as ET
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"
log_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\audit_all_sheets_output.txt"

def get_sheet_names(z):
    namelist = z.namelist()
    workbook_xml = ""
    workbook_rels = ""
    if "xl/workbook.xml" in namelist:
        workbook_xml = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
    if "xl/_rels/workbook.xml.rels" in namelist:
        workbook_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="ignore")
        
    rel_to_sheet = {}
    if workbook_rels:
        try:
            root = ET.fromstring(re.sub(r' xmlns="[^"]+"', '', workbook_rels))
            for rel in root.findall('.//Relationship'):
                rel_id = rel.get('Id')
                target = rel.get('Target')
                rel_to_sheet[rel_id] = target
        except Exception as e:
            pass
            
    sheet_mapping = {}
    if workbook_xml:
        try:
            workbook_clean = re.sub(r' xmlns="[^"]+"', '', workbook_xml)
            workbook_clean = re.sub(r' xmlns:[a-zA-Z0-9]+="[^"]+"', '', workbook_clean)
            workbook_clean = re.sub(r' [a-zA-Z0-9]+:[a-zA-Z0-9]+="[^"]+"', '', workbook_clean)
            workbook_clean = re.sub(r'r:id', 'id', workbook_clean)
            root = ET.fromstring(workbook_clean)
            for sheet in root.findall('.//sheet'):
                name = sheet.get('name')
                r_id = sheet.get('id')
                rel_path = rel_to_sheet.get(r_id)
                if rel_path:
                    full_path = "xl/" + rel_path if not rel_path.startswith("xl/") else rel_path
                    sheet_mapping[full_path] = name
        except Exception as e:
            # Fallback using regex
            matches = re.findall(r'<sheet name="([^"]+)"[^>]*id="([^"]+)"', workbook_xml)
            for name, r_id in matches:
                rel_path = rel_to_sheet.get(r_id)
                if rel_path:
                    full_path = "xl/" + rel_path if not rel_path.startswith("xl/") else rel_path
                    sheet_mapping[full_path] = name
    return sheet_mapping

def col_to_num(col_str):
    num = 0
    for char in col_str:
        if 'A' <= char <= 'Z':
            num = num * 26 + (ord(char) - ord('A') + 1)
    return num

def parse_cell_ref(ref):
    m = re.match(r'^([A-Z]+)([0-9]+)$', ref)
    if m:
        return col_to_num(m.group(1)), int(m.group(2))
    return 0, 0

def analyze_sheet(z, sheet_file, sheet_name):
    try:
        sheet_data = z.read(sheet_file).decode("utf-8", errors="ignore")
    except Exception as e:
        return {"error": str(e)}

    # Get declared dimension
    dim_match = re.search(r'<dimension ref="([^"]+)"', sheet_data)
    declared_dim = dim_match.group(1) if dim_match else "None"
    
    # Cells with row and col references
    # Cell match pattern
    cell_matches = re.finditer(r'<c r="([A-Z]+[0-9]+)"([^>]*)>(.*?)</c>|<c r="([A-Z]+[0-9]+)"([^>]*)\/>', sheet_data, re.DOTALL)
    
    total_cells = 0
    empty_styled_cells = 0
    ref_errors_in_formulas = 0
    ref_errors_in_values = 0
    has_formula_count = 0
    
    max_row_with_data = 0
    max_col_with_data = 0
    max_row_total = 0
    max_col_total = 0
    
    for m in cell_matches:
        total_cells += 1
        ref = m.group(1) if m.group(1) else m.group(4)
        attrs = m.group(2) if m.group(1) else m.group(5)
        inner = m.group(3) if m.group(1) else ""
        
        col_idx, row_idx = parse_cell_ref(ref)
        if row_idx > max_row_total:
            max_row_total = row_idx
        if col_idx > max_col_total:
            max_col_total = col_idx
            
        has_val = False
        has_f = False
        
        if inner:
            if '<v>' in inner:
                has_val = True
                # Check for error values like #REF!, #VALUE!, #N/A, #DIV/0!
                v_val = re.search(r'<v>(.*?)</v>', inner)
                if v_val and v_val.group(1).startswith('#'):
                    ref_errors_in_values += 1
            if '<f' in inner:
                has_f = True
                has_formula_count += 1
                if '#REF!' in inner or '#REF' in inner:
                    ref_errors_in_formulas += 1
            if '<t>' in inner or '<is>' in inner:
                has_val = True
                
        if has_val or has_f:
            if row_idx > max_row_with_data:
                max_row_with_data = row_idx
            if col_idx > max_col_with_data:
                max_col_with_data = col_idx
        else:
            if ' s="' in attrs:
                empty_styled_cells += 1
                
    # Also find if there is an autofilter and check if its range is larger than the data range
    autofilter_match = re.search(r'<autoFilter ref="([^"]+)"', sheet_data)
    autofilter_range = autofilter_match.group(1) if autofilter_match else "None"
    
    return {
        "sheet_name": sheet_name,
        "sheet_file": sheet_file,
        "declared_dim": declared_dim,
        "total_cells": total_cells,
        "empty_styled_cells": empty_styled_cells,
        "has_formula_count": has_formula_count,
        "ref_errors_in_formulas": ref_errors_in_formulas,
        "ref_errors_in_values": ref_errors_in_values,
        "max_row_total": max_row_total,
        "max_col_total": max_col_total,
        "max_row_with_data": max_row_with_data,
        "max_col_with_data": max_col_with_data,
        "autofilter_range": autofilter_range,
        "size_kb": len(sheet_data) / 1024
    }

def main():
    out = []
    
    def log(msg):
        out.append(msg)
        print(msg)

    with zipfile.ZipFile(file_path, 'r') as z:
        sheet_mapping = get_sheet_names(z)
        
        log(f"{'Sheet Name':<30} | {'File Size (KB)':<14} | {'Declared Dim':<15} | {'Max Row (Data)':<15} | {'Max Row (All)':<12} | {'Empty Styled':<12} | {'Formulas':<10} | {'Formula Errors'}")
        log("-" * 140)
        
        results = []
        for sfile in sorted(sheet_mapping.keys(), key=lambda x: int(re.search(r'sheet(\d+)', x).group(1)) if re.search(r'sheet(\d+)', x) else 0):
            sname = sheet_mapping[sfile]
            res = analyze_sheet(z, sfile, sname)
            results.append(res)
            
            if "error" in res:
                log(f"{sname:<30} | Error: {res['error']}")
            else:
                log(f"{res['sheet_name']:<30} | {res['size_kb']:>12.1f} KB | {res['declared_dim']:<15} | {res['max_row_with_data']:>15} | {res['max_row_total']:>12} | {res['empty_styled_cells']:>12} | {res['has_formula_count']:>10} | {res['ref_errors_in_formulas']:>14}")
                
        log("\n--- SHIFTED / GHOST RANGES DETECTED ---")
        for res in results:
            if "error" in res:
                continue
            # If max row with data is significantly smaller than max row total or declared dimension row
            dec_max_row = 0
            if ":" in res['declared_dim']:
                parts = res['declared_dim'].split(":")
                if len(parts) == 2:
                    _, dec_max_row = parse_cell_ref(parts[1])
            
            diff_dec = dec_max_row - res['max_row_with_data']
            diff_total = res['max_row_total'] - res['max_row_with_data']
            
            if diff_dec > 10 or diff_total > 10:
                log(f"Sheet '{res['sheet_name']}' has potential GHOST RANGES:")
                log(f"  - Max Row with Data/Formula: {res['max_row_with_data']}")
                log(f"  - Declared Max Row: {dec_max_row} (Difference: {diff_dec})")
                log(f"  - Max Row present in XML: {res['max_row_total']} (Difference: {diff_total})")
                log(f"  - Empty styled cells: {res['empty_styled_cells']}")
                if res['autofilter_range'] != "None":
                    log(f"  - AutoFilter Range: {res['autofilter_range']}")
                    
        log("\n--- FORMULA ERROR REPORT ---")
        for res in results:
            if "error" in res:
                continue
            if res['ref_errors_in_formulas'] > 0 or res['ref_errors_in_values'] > 0:
                log(f"Sheet '{res['sheet_name']}' has errors:")
                log(f"  - Formulas containing '#REF!': {res['ref_errors_in_formulas']}")
                log(f"  - Cell values containing errors: {res['ref_errors_in_values']}")

    with open(log_path, "w", encoding="utf-8") as lf:
        lf.write("\n".join(out))
    print(f"\nSaved full log to {log_path}")

if __name__ == "__main__":
    main()
