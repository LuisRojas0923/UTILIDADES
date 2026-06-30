import os
import re
import base64
import xml.etree.ElementTree as ET

item6_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\item6_decoded.xml"

if not os.path.exists(item6_path):
    print("File not found")
    exit()

with open(item6_path, "r", encoding="utf-8") as f:
    content = f.read()

print(f"Decoded XML file size: {len(content)} characters")
print("First 1500 characters of XML:")
print(content[:1500])

# Let's search for tags using a simple regex to see what elements are present
tags = re.findall(r'<([A-Za-z0-9:]+)', content)
print("\nUnique tags found in XML:")
print(set(tags))

# Look for base64 strings (usually long strings with letters, numbers, +, /, sometimes ending with =)
# In a DataMashup, there is a base64 string inside some tag or we can find it.
# Let's see if there is any long text.
long_texts = re.findall(r'>([A-Za-z0-9+/=\r\n]{1000,})<', content)
print(f"\nFound {len(long_texts)} text blocks longer than 1000 chars.")
for idx, text in enumerate(long_texts):
    clean_text = re.sub(r'\s+', '', text)
    print(f"Block #{idx+1} length (cleaned): {len(clean_text)}")
    try:
        decoded = base64.b64decode(clean_text)
        print(f"  Decoded size: {len(decoded)} bytes")
        # Check if the decoded data contains a ZIP file header PK\x03\x04
        pk_idx = decoded.find(b'PK\x03\x04')
        if pk_idx != -1:
            print(f"  Found ZIP header inside decoded block at index {pk_idx}!")
            zip_out_path = f"C:\\Users\\amejoramiento6\\Desktop\\UTILIDADES\\scratch\\decoded_mashup_{idx+1}.zip"
            with open(zip_out_path, "wb") as fz:
                fz.write(decoded[pk_idx:])
            print(f"  Wrote ZIP to {zip_out_path}")
            # List contents
            import zipfile
            with zipfile.ZipFile(zip_out_path, 'r') as mz:
                print(f"  ZIP contents: {mz.namelist()}")
                for mname in mz.namelist():
                    if "Section1.section" in mname:
                        sec_content = mz.read(mname).decode('utf-8', errors='ignore')
                        # save to clean text
                        pq_out = f"C:\\Users\\amejoramiento6\\Desktop\\UTILIDADES\\scratch\\power_queries_extracted.txt"
                        with open(pq_out, "w", encoding="utf-8") as f_pq:
                            f_pq.write(sec_content)
                        print(f"  EXTRACTED POWER QUERY FORMULAS TO {pq_out}")
    except Exception as e:
        print(f"  Failed to decode block: {e}")
