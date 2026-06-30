import zipfile
import re

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"

def peek_rows():
    row_re = re.compile(r'<row r="(\d+)"[^>]*>(.*?)</row>', re.DOTALL)
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            with zip_ref.open('xl/worksheets/sheet1.xml') as sheet:
                head = sheet.read(2 * 1024 * 1024).decode('utf-8', errors='ignore')
                print("--- PEEKING ROWS 1-100 ---")
                rows = row_re.findall(head)
                for r, content in rows[:100]:
                    # Print row number and a bit of content
                    print(f"Row {r}: {content[:80]}...")
    except Exception as e:
        print(f"Error: {e}")

peek_rows()
