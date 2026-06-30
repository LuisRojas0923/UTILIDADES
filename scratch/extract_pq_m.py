import zipfile

zip_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\temp_zip_0.zip"
out_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\power_queries_extracted_ok.txt"

with zipfile.ZipFile(zip_path, 'r') as z:
    for name in z.namelist():
        if "Section1.m" in name:
            m_code = z.read(name).decode('utf-8', errors='ignore')
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(m_code)
            print(f"Extracted Formulas/Section1.m to {out_path} ({len(m_code)} bytes)")
