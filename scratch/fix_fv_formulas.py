import zipfile
import os
import re

opt_file = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED.xlsm"
temp_file = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED_2.xlsm"

def fix_fv_formulas():
    if not os.path.exists(opt_file):
        print("Optimized file not found")
        return False
        
    print(f"Fixing FV formulas in {opt_file}")
    
    replacements = 0
    with zipfile.ZipFile(opt_file, 'r') as zin:
        with zipfile.ZipFile(temp_file, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                filename = item.filename
                data = zin.read(filename)
                
                if filename == "xl/worksheets/sheet2.xml":
                    content = data.decode("utf-8", errors="ignore")
                    
                    # Target: IF(QUINCE.AUDITAR=BNF.Q1,#REF!,
                    # We want to replace it with: IF(QUINCE.AUDITAR=BNF.Q1,Descuentos_Consolidados[BENEFICIAR.1Q APORTE],
                    target = "IF(QUINCE.AUDITAR=BNF.Q1,#REF!,"
                    replacement = "IF(QUINCE.AUDITAR=BNF.Q1,Descuentos_Consolidados[BENEFICIAR.1Q APORTE],"
                    
                    matches = len(re.findall(re.escape(target), content))
                    if matches > 0:
                        content = content.replace(target, replacement)
                        replacements = matches
                        print(f"Replaced {matches} occurrences of broken IF(#REF) reference.")
                    else:
                        print("Target IF(#REF) reference not found.")
                        
                    data = content.encode("utf-8")
                
                zout.writestr(filename, data)
                
    # Replace optimized file with the new one
    os.remove(opt_file)
    os.rename(temp_file, opt_file)
    print(f"Completed! Replaced: {replacements}")
    return True

if __name__ == "__main__":
    fix_fv_formulas()
