import zipfile
import binascii

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"

def hex_dump(data, start_offset=0):
    lines = []
    for i in range(0, len(data), 16):
        chunk = data[i:i+16]
        hex_str = " ".join(f"{b:02x}" for b in chunk)
        ascii_str = "".join(chr(b) if 32 <= b <= 126 else "." for b in chunk)
        lines.append(f"{start_offset + i:08x}: {hex_str:<48} |{ascii_str}|")
    return "\n".join(lines)

print("--- FILE START (First 256 bytes) ---")
with open(file_path, "rb") as f:
    start_data = f.read(256)
    print(hex_dump(start_data))

print("\n--- FILE END (Last 256 bytes) ---")
with open(file_path, "rb") as f:
    f.seek(0, 2)
    size = f.tell()
    f.seek(max(0, size - 256))
    end_data = f.read(256)
    print(hex_dump(end_data, start_offset=max(0, size - 256)))

print("\n--- INTERNAL SHEET1.XML PREVIEW ---")
try:
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        with zip_ref.open('xl/worksheets/sheet1.xml') as sheet:
            preview = sheet.read(2000)
            print(preview.decode('utf-8', errors='ignore'))
except Exception as e:
    print(f"Error reading internal file: {e}")
