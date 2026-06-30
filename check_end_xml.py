import zipfile

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"

try:
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        with zip_ref.open('xl/worksheets/sheet1.xml') as sheet:
            # The file is 400MB uncompressed. 
            # Let's seek to near the end. 
            # ZipExtFile doesn't support seek well for compressed files if we want to skip.
            # But we can read in chunks.
            
            chunk_size = 1024 * 1024 # 1MB
            last_chunk = b""
            while True:
                data = sheet.read(chunk_size)
                if not data:
                    break
                last_chunk = data
            
            print("--- LAST 2000 CHARACTERS OF SHEET1.XML ---")
            print(last_chunk[-2000:].decode('utf-8', errors='ignore'))
            
except Exception as e:
    print(f"Error: {e}")
