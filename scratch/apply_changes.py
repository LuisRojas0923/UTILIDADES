import os
import shutil

src = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"
backup = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_BACKUP.xlsm"
opt = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2_OPTIMIZED.xlsm"

def apply():
    if not os.path.exists(opt):
        print("Optimized file not found")
        return False
        
    print(f"Backing up original to: {backup}")
    shutil.copy2(src, backup)
    
    print(f"Overwriting original with optimized file...")
    shutil.move(opt, src)
    
    print("Changes applied successfully!")
    return True

if __name__ == "__main__":
    apply()
