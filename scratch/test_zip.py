import zipfile

zip_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\decoded_mashup_1.zip"

try:
    with zipfile.ZipFile(zip_path, 'r') as z:
        print(f"File list: {z.namelist()}")
        print(f"Infolist: {z.infolist()}")
        for info in z.infolist():
            print(f"File: {info.filename}, Size: {info.file_size}")
except Exception as e:
    print(f"Error: {e}")
