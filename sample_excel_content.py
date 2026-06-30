import zipfile
import re

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"

def sample_rows():
    row_re = re.compile(r'<row r="(\d+)"[^>]*>(.*?)</row>', re.DOTALL)
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            with zip_ref.open('xl/worksheets/sheet1.xml') as sheet:
                # Read first 1MB
                head = sheet.read(1024 * 1024).decode('utf-8', errors='ignore')
                print("--- FIRST ROWS ---")
                for r, content in row_re.findall(head)[:5]:
                    print(f"Row {r}: {content[:100]}...")
                
                # Sample middle (this is hard without full read, but let's try to skip)
                # Since it's 400MB, let's read around 200MB mark
                # We can't seek in ZipExtFile easily for compression, but we can read and discard
                print("\n--- SAMPLING AT VARIOUS POINTS ---")
                
                # Reset or just continue reading
                # Let's read in 50MB chunks and report the first row found in each
                for i in range(1, 9):
                    sheet.read(50 * 1024 * 1024) # Skip 50MB
                    chunk = sheet.read(1024 * 1024).decode('utf-8', errors='ignore')
                    match = row_re.search(chunk)
                    if match:
                        r, content = match.groups()
                        print(f"Around {i*50}MB: Found Row {r} -> Content: {content[:150]}...")

    except Exception as e:
        print(f"Error: {e}")

sample_rows()
