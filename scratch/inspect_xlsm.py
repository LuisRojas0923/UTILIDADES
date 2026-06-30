import zipfile
import os
import xml.etree.ElementTree as ET
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"

def analyze_xlsm(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    print(f"Analyzing: {os.path.basename(file_path)}")
    print(f"File size: {os.path.getsize(file_path) / (1024*1024):.2f} MB")
    
    with zipfile.ZipFile(file_path, 'r') as z:
        namelist = z.namelist()
        
        # 1. Top files by size
        print("\n--- TOP 15 LARGEST INTERNAL FILES ---")
        info_list = z.infolist()
        info_list.sort(key=lambda x: x.file_size, reverse=True)
        for info in info_list[:15]:
            print(f"{info.filename:<50} | {info.file_size:>12,} bytes (Unzipped) | {info.compress_size:>12,} bytes (Zipped)")
            
        # 2. Check for VBA and Connections
        print("\n--- MODULES & EXTERNAL FEATURES ---")
        vba_exists = any("vbaProject.bin" in name for name in namelist)
        connections_exists = any("connections.xml" in name for name in namelist)
        query_tables = [name for name in namelist if "queryTable" in name]
        pivot_tables = [name for name in namelist if "pivotTable" in name]
        calc_chain_exists = any("calcChain.xml" in name for name in namelist)
        
        print(f"VBA Macro Project (vbaProject.bin): {'YES' if vba_exists else 'NO'}")
        print(f"External Connections (connections.xml): {'YES' if connections_exists else 'NO'}")
        print(f"Power Query / Query Tables: {len(query_tables)} found")
        print(f"Pivot Tables: {len(pivot_tables)} found")
        print(f"Calculation Chain (calcChain.xml): {'YES' if calc_chain_exists else 'NO'}")
        
        # 3. Read sheet names mapping
        print("\n--- WORKSHEETS DETAILS ---")
        sheets_info = {}
        workbook_xml = ""
        workbook_rels = ""
        
        if "xl/workbook.xml" in namelist:
            workbook_xml = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
        if "xl/_rels/workbook.xml.rels" in namelist:
            workbook_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="ignore")
            
        # Parse rels
        rel_to_sheet = {}
        if workbook_rels:
            try:
                root = ET.fromstring(workbook_rels)
                # handle namespaces
                ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rel_id = rel.get('Id')
                    target = rel.get('Target')
                    rel_to_sheet[rel_id] = target
            except Exception as e:
                print(f"Error parsing workbook.xml.rels: {e}")
                
        # Parse workbook to map sheet names to xml files
        sheet_mapping = {}
        if workbook_xml:
            try:
                # remove namespace to make parsing easier with regex or basic ET
                workbook_clean = re.sub(r' xmlns="[^"]+"', '', workbook_xml)
                root = ET.fromstring(workbook_clean)
                for sheet in root.findall('.//sheet'):
                    name = sheet.get('name')
                    r_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    rel_path = rel_to_sheet.get(r_id)
                    if rel_path:
                        # normalize path
                        full_path = "xl/" + rel_path if not rel_path.startswith("xl/") else rel_path
                        sheet_mapping[full_path] = name
            except Exception as e:
                print(f"Error parsing workbook.xml: {e}")
                # Try regex fallback
                matches = re.findall(r'<sheet name="([^"]+)"[^>]*id="([^"]+)"', workbook_xml)
                for name, r_id in matches:
                    rel_path = rel_to_sheet.get(r_id)
                    if rel_path:
                        full_path = "xl/" + rel_path if not rel_path.startswith("xl/") else rel_path
                        sheet_mapping[full_path] = name

        # 4. Analyze each sheet
        print(f"{'Sheet Name':<30} | {'Internal File':<25} | {'Dimension':<15} | {'Size (MB)':<10} | {'Estimated Cells'}")
        print("-" * 100)
        
        for name in namelist:
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                sheet_data = z.read(name)
                sheet_size = len(sheet_data) / (1024*1024)
                
                # Get dimension
                dim_match = re.search(r'<dimension ref="([^"]+)"', sheet_data.decode("utf-8", errors="ignore"))
                dim = dim_match.group(1) if dim_match else "Unknown"
                
                # Count cells (c tags)
                cell_count = len(re.findall(r'<c ', sheet_data.decode("utf-8", errors="ignore")))
                # Count formulas (f tags)
                formula_count = len(re.findall(r'<f[ >]', sheet_data.decode("utf-8", errors="ignore")))
                
                sheet_name = sheet_mapping.get(name, "Unknown Sheet Name")
                print(f"{sheet_name:<30} | {name:<25} | {dim:<15} | {sheet_size:>10.2f} MB | {cell_count:,} cells ({formula_count:,} formulas)")

if __name__ == "__main__":
    analyze_xlsm(file_path)
