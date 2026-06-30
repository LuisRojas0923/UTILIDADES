import zipfile
import os

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"

with zipfile.ZipFile(file_path, 'r') as z:
    for name in z.namelist():
        if "customXml" in name:
            data = z.read(name)
            print(f"{name}: {len(data)} bytes | Preview: {data[:120]}")
