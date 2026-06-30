import zipfile
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED.xlsm"

with zipfile.ZipFile(file_path, 'r') as z:
    sheet_data = z.read("xl/worksheets/sheet2.xml").decode("utf-8", errors="ignore")
    
    # Let's inspect column FR on row 195 and row 196
    for row in [195, 196, 197]:
        # Search for cell ref FR+row or FS+row or FT+row, let's look for matching row r="row" inside sheetData
        # and print the cells in that row.
        # We can find the row block: <row r="row" ...> ... </row>
        row_block_match = re.search(r'<row r="' + str(row) + r'"[^>]*>(.*?)</row>', sheet_data, re.DOTALL)
        if row_block_match:
            row_content = row_block_match.group(1)
            # Find cells containing #REF! in this row
            cells = re.finditer(r'<c r="([A-Z]+[0-9]+)"([^>]*)>(.*?)</c>', row_content, re.DOTALL)
            print(f"\n--- Row {row} Cells with #REF! ---")
            for c in cells:
                ref = c.group(1)
                inner = c.group(3)
                if '#REF!' in inner:
                    f_match = re.search(r'<f[^>]*>(.*?)</f>', inner)
                    formula = f_match.group(1) if f_match else "No formula (just value)"
                    print(f"  Cell {ref}: Formula={formula[:120]}")
        else:
            print(f"Row {row} not found")
