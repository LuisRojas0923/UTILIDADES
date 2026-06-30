import zipfile
import os
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"
log_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\vba_strings_output.txt"

def extract_strings(data, min_len=4):
    # ASCII strings
    ascii_re = re.compile(br'[ -~]{' + str(min_len).encode() + br',}')
    ascii_strings = [s.decode('ascii', errors='ignore') for s in ascii_re.findall(data)]
    
    # UTF-16 strings (unicode)
    unicode_re = re.compile(br'(?:[ -~]\x00){' + str(min_len).encode() + br',}')
    unicode_strings = [s.decode('utf-16le', errors='ignore') for s in unicode_re.findall(data)]
    
    return ascii_strings + unicode_strings

def main():
    if not os.path.exists(file_path):
        print("File not found")
        return
        
    out = []
    def log(msg):
        out.append(msg)
        print(msg)

    with zipfile.ZipFile(file_path, 'r') as z:
        namelist = z.namelist()
        vba_file = "xl/vbaProject.bin"
        if vba_file not in namelist:
            log("No xl/vbaProject.bin found in the xlsm file.")
            return
            
        data = z.read(vba_file)
        log(f"Extracting strings from {vba_file} ({len(data)} bytes)...")
        
        strings = extract_strings(data, min_len=5)
        log(f"Extracted {len(strings)} strings.")
        
        # Filter strings of interest: Sub, Function, Select, Insert, Update, Delete, Connection, http, Sheet, Table, etc.
        vba_keywords = [
            r'\bSub\b', r'\bFunction\b', r'\bDim\b', r'\bCall\b', r'\bWorkbook_Open\b', 
            r'\bActiveSheet\b', r'\bRange\b', r'\bCells\b', r'\bSelection\b', r'\bSelect\b',
            r'\bSQL\b', r'\bConnection\b', r'\bProvider\b', r'\bServer\b', r'\bDatabase\b',
            r'\bSheets\b', r'\bWorksheets\b', r'\bMsgBox\b', r'\bScreenUpdating\b',
            r'\bCalculation\b', r'\bCalculate\b', r'\bRefreshAll\b', r'\bRefresh\b',
            r'\bQueryTable\b', r'\bPowerQuery\b'
        ]
        
        compiled_kw = [re.compile(kw, re.IGNORECASE) for kw in vba_keywords]
        
        interesting_strings = []
        for s in strings:
            s_clean = s.strip()
            if not s_clean:
                continue
            for r_kw in compiled_kw:
                if r_kw.search(s_clean):
                    interesting_strings.append(s_clean)
                    break
                    
        log(f"Found {len(interesting_strings)} interesting/keyword-matching strings:")
        
        # Sort and unique
        seen = set()
        unique_interesting = []
        for s in interesting_strings:
            if s.lower() not in seen:
                seen.add(s.lower())
                unique_interesting.append(s)
                
        log(f"Found {len(unique_interesting)} unique interesting strings:")
        for s in unique_interesting[:150]: # top 150 unique strings
            log(f"  - {s}")
            
    with open(log_path, "w", encoding="utf-8") as lf:
        lf.write("\n".join(out) + "\n\n=== ALL STRINGS ===\n" + "\n".join(set(strings)))
    print(f"\nSaved full strings to {log_path}")

if __name__ == "__main__":
    main()
