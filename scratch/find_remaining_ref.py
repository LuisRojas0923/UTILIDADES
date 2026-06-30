import zipfile
import re

opt_file = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED.xlsm"

with zipfile.ZipFile(opt_file, 'r') as z:
    sheet2_data = z.read("xl/worksheets/sheet2.xml").decode("utf-8", errors="ignore")
    
    # Find all occurrences of #REF! in the entire XML
    matches = re.finditer(r'#REF!', sheet2_data)
    print("All #REF! occurrences in sheet2:")
    for i, m in enumerate(matches):
        idx = m.start()
        # print surrounding context
        context = sheet2_data[max(0, idx-100):min(len(sheet2_data), idx+100)]
        # clean xml formatting a bit for reading
        context_clean = context.replace('&quot;', '"').replace('&lt;', '<').replace('&gt;', '>')
        print(f"Occurrence #{i+1} (index {idx}):\n  ... {context_clean} ...\n")
