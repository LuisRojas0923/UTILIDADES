import zipfile
import re

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"

def find_break_point():
    row_re = re.compile(r'<row r="(\d+)"[^>]*>(.*?)</row>', re.DOTALL)
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            with zip_ref.open('xl/worksheets/sheet1.xml') as sheet:
                # Read 5MB to be sure
                head = sheet.read(5 * 1024 * 1024).decode('utf-8', errors='ignore')
                rows = row_re.findall(head)
                
                print(f"Total rows in head: {len(rows)}")
                for r, content in rows:
                    r_int = int(r)
                    # Check if row has column B (r="B...")
                    if f'r="B{r}"' not in content:
                        print(f"Row {r} is the first row MISSING column B.")
                        # Peek at content
                        print(f"Content: {content[:200]}")
                        return r_int
                
                return None
    except Exception as e:
        print(f"Error: {e}")
        return None

break_row = find_break_point()
print(f"Break row: {break_row}")
