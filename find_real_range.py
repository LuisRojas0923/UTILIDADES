import zipfile
import re

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"

def find_last_real_row():
    last_row_with_data = 0
    # Regular expression to find row index and check if it has a value <v> or <s> with data
    # In Excel XML, <c r="A1" s="1"><v>Value</v></c>
    # Ghost rows often look like <c r="A100" s="1"/> (no value)
    
    row_re = re.compile(r'<row r="(\d+)"')
    val_re = re.compile(r'<v>|<t>|<is>') # Tags that indicate actual content
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            with zip_ref.open('xl/worksheets/sheet1.xml') as sheet:
                # We'll read in large chunks and look for the highest row number that has a value
                chunk_size = 1024 * 1024 * 5 # 5MB chunks
                current_row = 0
                
                while True:
                    chunk = sheet.read(chunk_size).decode('utf-8', errors='ignore')
                    if not chunk:
                        break
                    
                    # Find all rows in this chunk
                    rows = row_re.findall(chunk)
                    if rows:
                        # Check if any row in this chunk has values
                        # This is an approximation: if the chunk has <v>, we assume the rows in it have data
                        if val_re.search(chunk):
                            last_row_with_data = max(last_row_with_data, int(rows[-1]))
                        
                return last_row_with_data
    except Exception as e:
        print(f"Error: {e}")
        return None

last_row = find_last_real_row()
print(f"Estimated last row with actual data: {last_row}")
