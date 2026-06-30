import zipfile
import re

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026_FIXED.xlsx"

def verify():
    row_re = re.compile(r'<row r="(\d+)"')
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            with zip_ref.open('xl/worksheets/sheet1.xml') as sheet:
                content = sheet.read().decode('utf-8', errors='ignore')
                rows = row_re.findall(content)
                print(f"Total rows in fixed file: {len(rows)}")
                if rows:
                    print(f"First row: {rows[0]}, Last row: {rows[-1]}")
                
                # Check for dimension
                dim_match = re.search(r'dimension ref="([^"]+)"', content)
                if dim_match:
                    print(f"Dimension: {dim_match.group(1)}")
    except Exception as e:
        print(f"Error: {e}")

verify()
