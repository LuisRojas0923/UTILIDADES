import zipfile
import os
import xml.etree.ElementTree as ET
import re
from collections import Counter

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"
log_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\inspect_details_output.txt"

def get_xml_root(zip_file, filename):
    try:
        content = zip_file.read(filename)
        content_clean = re.sub(r' xmlns="[^"]+"', '', content.decode("utf-8", errors="ignore"))
        content_clean = re.sub(r' xmlns:[a-zA-Z0-9]+="[^"]+"', '', content_clean)
        content_clean = re.sub(r' [a-zA-Z0-9]+:[a-zA-Z0-9]+="[^"]+"', '', content_clean)
        content_clean = re.sub(r'</?[a-zA-Z0-9]+:', lambda m: '</' if m.group().startswith('</') else '<', content_clean)
        return ET.fromstring(content_clean)
    except Exception as e:
        return None

def main():
    out = []
    
    def log(msg):
        out.append(msg)
        print(msg)

    with zipfile.ZipFile(file_path, 'r') as z:
        namelist = z.namelist()
        
        # Connections
        log("\n==================================================")
        log("ANALYSIS OF DATABASE CONNECTIONS & POWER QUERIES")
        log("==================================================")
        
        conn_file = "xl/connections.xml"
        if conn_file in namelist:
            root = get_xml_root(z, conn_file)
            if root is not None:
                connections = root.findall('.//connection')
                log(f"Found {len(connections)} connections in connections.xml:")
                for conn in connections:
                    conn_id = conn.get('id')
                    name = conn.get('name')
                    conn_type = conn.get('type')
                    description = conn.get('description', '')
                    log(f"\nConnection #{conn_id}: Name='{name}', Type='{conn_type}'")
                    if description:
                        log(f"  Description: {description}")
                    db_pr = conn.find('.//dbPr')
                    if db_pr is not None:
                        log(f"  Connection String: {db_pr.get('connection')}")
                        command = db_pr.get('command')
                        if command:
                            log(f"  SQL Query / Command: {command}")
        
        # Query Tables
        qtable_files = [name for name in namelist if name.startswith("xl/queryTables/queryTable")]
        if qtable_files:
            log(f"\nFound {len(qtable_files)} Query Tables:")
            for qtf in qtable_files:
                qt_root = get_xml_root(z, qtf)
                if qt_root is not None:
                    qt = qt_root.find('.//queryTable')
                    if qt is not None:
                        log(f"  - {qtf}: Name='{qt.get('name')}', connectionId='{qt.get('connectionId')}', rowNumbers='{qt.get('rowNumbers')}'")
                        
        # Power Query formulas in customXml
        custom_xml_files = [name for name in namelist if name.startswith("customXml/item") and name.endswith(".xml")]
        for cxf in custom_xml_files:
            content = z.read(cxf).decode("utf-8", errors="ignore")
            if "Section1" in content or "Expression" in content or "shared" in content:
                log(f"\nPower Query Formula Container detected in: {cxf}")
                m_code = re.findall(r'formula="([^"]+)"', content)
                if m_code:
                    log("  Power Query Formulas (M Code) found:")
                    for formula in m_code:
                        formula_decoded = formula.replace('&quot;', '"').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
                        log(f"    - {formula_decoded}")
                else:
                    # Look for other M code sections
                    # Let's extract everything inside tag data or expression if it's there
                    m_expressions = re.findall(r'<[a-zA-Z0-9:]*expression[^>]*>(.*?)</[a-zA-Z0-9:]*expression>', content, re.DOTALL)
                    if m_expressions:
                        for exp in m_expressions:
                            log(f"    - Expression: {exp[:300]}...")
                    else:
                        snippet = content[:2000].replace('&quot;', '"').replace('&lt;', '<').replace('&gt;', '>')
                        log(f"  Content Snippet (first 2000 chars):\n{snippet}")

        # Analyze sheets: AUDITORIA NOMINA (sheet2) and ESTABLECIMIENTO_SOLID (sheet7)
        for sfile, sname in [("xl/worksheets/sheet2.xml", "AUDITORIA NOMINA"), ("xl/worksheets/sheet7.xml", "ESTABLECIMIENTO_SOLID")]:
            log("\n==================================================")
            log(f"DETAILED ANALYSIS FOR SHEET: {sname} ({sfile})")
            log("==================================================")
            
            if sfile not in namelist:
                log(f"Sheet file {sfile} not found.")
                continue
                
            sheet_data = z.read(sfile).decode("utf-8", errors="ignore")
            all_cells = re.findall(r'<c r="([A-Z]+[0-9]+)"([^>]*)>(.*?)</c>|<c r="([A-Z]+[0-9]+)"([^>]*)\/>', sheet_data)
            
            total_cells = 0
            empty_styled_cells = 0
            value_cells = 0
            formula_cells = 0
            formulas = []
            
            for match in all_cells:
                total_cells += 1
                r = match[0] if match[0] else match[3]
                attrs = match[1] if match[0] else match[4]
                inner = match[2] if match[0] else ""
                has_style = ' s="' in attrs
                
                if not inner:
                    if has_style:
                        empty_styled_cells += 1
                else:
                    has_v = '<v>' in inner
                    has_f = '<f' in inner
                    has_t = ' t="' in attrs or '<is>' in inner
                    
                    if has_f:
                        formula_cells += 1
                        f_match = re.search(r'<f[^>]*>(.*?)</f>', inner)
                        if f_match:
                            formulas.append(f_match.group(1))
                        else:
                            formulas.append("SHARED_FORMULA")
                    elif has_v or has_t:
                        value_cells += 1
                    else:
                        if has_style:
                            empty_styled_cells += 1
                            
            log(f"Total cells processed: {total_cells:,}")
            log(f"  - Cells with data (values/text): {value_cells:,}")
            log(f"  - Cells with formulas: {formula_cells:,}")
            log(f"  - Empty cells with formatting (styled but no content): {empty_styled_cells:,}")
            if total_cells > 0:
                log(f"  - Format-only overhead: {empty_styled_cells / total_cells * 100:.1f}% of cells are empty styled")
                
            if formulas:
                log(f"\nAnalyzing Formulas ({len(formulas)} formulas):")
                patterns = []
                for f in formulas:
                    f_pat = re.sub(r'\b[A-Z]+[0-9]+\b', 'REF', f)
                    f_pat = re.sub(r'"[^"]*"', '"STR"', f_pat)
                    f_pat = re.sub(r'\b\d+(?:\.\d+)?\b', 'NUM', f_pat)
                    patterns.append(f_pat)
                    
                counter = Counter(patterns)
                log("\nTop 10 Formula Patterns:")
                for pat, count in counter.most_common(10):
                    log(f"  - [{count:>5} occurrences]: {pat}")
                    
                expensive_funcs = ['VLOOKUP', 'HLOOKUP', 'INDIRECT', 'OFFSET', 'SUMIFS', 'COUNTIFS', 'AVERAGEIFS', 'MATCH', 'INDEX', 'XLOOKUP']
                expensive_counts = Counter()
                for f in formulas:
                    for func in expensive_funcs:
                        if func in f.upper():
                            expensive_counts[func] += 1
                            
                log("\nExpensive/Volatile Function Usage Counts:")
                for func, count in expensive_counts.items():
                    log(f"  - {func}: {count} occurrences")
                    
                log("\nSample Unique/Complex Formulas:")
                unique_formulas = list(set(formulas))
                unique_formulas.sort(key=len, reverse=True)
                for uf in unique_formulas[:25]:
                    log(f"  - {uf}")
                    
    with open(log_path, "w", encoding="utf-8") as lf:
        lf.write("\n".join(out))
    print(f"\nSaved full log to {log_path}")

if __name__ == "__main__":
    main()
