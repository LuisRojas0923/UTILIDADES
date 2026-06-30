import zipfile
import re

opt_file = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED.xlsm"

def verify():
    print(f"Verifying: {opt_file}")
    
    try:
        with zipfile.ZipFile(opt_file, 'r') as z:
            namelist = z.namelist()
            print("  - ZIP structure is valid.")
            
            # Check sheet2 for #REF!
            sheet2_data = z.read("xl/worksheets/sheet2.xml").decode("utf-8", errors="ignore")
            # We had #REF! in sheet2. Are there any left in <f> tags?
            ref_formulas = re.findall(r'<f[^>]*>[^<]*#REF![^<]*</f>', sheet2_data)
            print(f"  - Sheet 2: Remaining #REF! formulas: {len(ref_formulas)}")
            if len(ref_formulas) > 0:
                print(f"    Sample remaining error formulas: {ref_formulas[:5]}")
                
            # Check sheet18 for cell N311
            sheet18_data = z.read("xl/worksheets/sheet18.xml").decode("utf-8", errors="ignore")
            has_n311 = 'r="N311"' in sheet18_data
            print(f"  - Sheet 18: Contains cell N311: {has_n311}")
            
            # Check sheet9 dimension and row count
            sheet9_data = z.read("xl/worksheets/sheet9.xml").decode("utf-8", errors="ignore")
            dim_match = re.search(r'<dimension ref="([^"]+)"', sheet9_data)
            dim = dim_match.group(1) if dim_match else "None"
            print(f"  - Sheet 9: Dimension = {dim}")
            
            # Count row tags
            row_tags = len(re.findall(r'<row ', sheet9_data))
            print(f"  - Sheet 9: Number of row tags: {row_tags}")
            
    except Exception as e:
        print(f"Verification failed: {e}")

if __name__ == "__main__":
    verify()
