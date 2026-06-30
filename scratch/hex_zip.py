import zipfile
import base64
import os
import re

item6_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\item6_decoded.xml"

with open(item6_path, "r", encoding="utf-8") as f:
    content = f.read()

long_texts = re.findall(r'>([A-Za-z0-9+/=\r\n]{1000,})<', content)
clean_text = re.sub(r'\s+', '', long_texts[0])
decoded = base64.b64decode(clean_text)

print(f"Decoded data length: {len(decoded)} bytes")

# Find all PK\x03\x04 occurrences
offsets = []
idx = 0
while True:
    pk_idx = decoded.find(b'PK\x03\x04', idx)
    if pk_idx == -1:
        break
    offsets.append(pk_idx)
    idx = pk_idx + 1

print(f"All PK\\x03\\x04 offsets: {offsets}")

# Try to extract ZIP files starting at each offset
for i, offset in enumerate(offsets):
    print(f"\nTrying offset {offset}:")
    zip_candidate = decoded[offset:]
    temp_zip_path = f"C:\\Users\\amejoramiento6\\Desktop\\UTILIDADES\\scratch\\temp_zip_{i}.zip"
    
    # We need to find where the ZIP ends.
    # In OLE DataMashup, there is a length header before the ZIP file.
    # The header before the ZIP file (at index 8 in decoded data, or actually index 4 to 7) is a 4-byte little-endian length.
    # Let's check the bytes before the PK\x03\x04
    if offset >= 4:
        length_bytes = decoded[offset-4:offset]
        length = int.from_bytes(length_bytes, byteorder='little')
        print(f"  Length header before ZIP: {length} bytes")
        # Slice only the length of the ZIP
        zip_candidate = decoded[offset:offset+length]
        
    with open(temp_zip_path, "wb") as fz:
        fz.write(zip_candidate)
        
    try:
        with zipfile.ZipFile(temp_zip_path, 'r') as z:
            names = z.namelist()
            print(f"  SUCCESS! Zip contains {len(names)} files:")
            for name in names:
                print(f"    - {name} ({len(z.read(name))} bytes)")
                if "Section1.section" in name:
                    # extract formulas
                    m_code = z.read(name).decode('utf-8', errors='ignore')
                    out_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\power_queries_extracted_ok.txt"
                    with open(out_path, "w", encoding="utf-8") as f_pq:
                        f_pq.write(m_code)
                    print(f"  Extracted Power Query code to {out_path} (length: {len(m_code)})")
    except Exception as e:
        print(f"  Failed: {e}")
