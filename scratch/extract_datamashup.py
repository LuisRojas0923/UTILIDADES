import zipfile
import os
import xml.etree.ElementTree as ET
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"
log_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\power_queries_clean.txt"

def decode_data(data):
    if data.startswith(b'\xff\xfe'):
        return data.decode('utf-16', errors='ignore')
    elif data.startswith(b'\xfe\xff'):
        return data.decode('utf-16be', errors='ignore')
    else:
        return data.decode('utf-8', errors='ignore')

def main():
    if not os.path.exists(file_path):
        print("File not found")
        return
        
    out = []
    def log(msg):
        out.append(msg)
        print(msg[:300]) # Print first 300 chars to terminal

    with zipfile.ZipFile(file_path, 'r') as z:
        # Check customXml/item6.xml
        name = "customXml/item6.xml"
        if name in z.namelist():
            raw_data = z.read(name)
            content = decode_data(raw_data)
            
            # The XML starts with <DataMashup ...>
            # Inside it might have binary mashup parts.
            # Power Query code is usually inside the mashup package.
            # In item6.xml, there is often a Base64-encoded or raw binary package.
            # Let's inspect what elements are in item6.xml.
            # We can use regex to extract everything that looks like M code or step definitions.
            log(f"Length of {name} content: {len(content)} characters")
            
            # Find all occurrences of let ... in in the decoded XML
            # Let's clean up any binary garbage first:
            # We can look for strings matching the pattern of M code:
            # e.g., 'let\r\n' or 'shared [A-Za-z0-9_ ]+ = let'
            
            # Let's write the whole XML content to a temporary text file so we can view it
            temp_xml_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\item6_decoded.xml"
            with open(temp_xml_path, "w", encoding="utf-8") as tf:
                tf.write(content)
            log(f"Wrote decoded xml to {temp_xml_path}")
            
            # Let's find Section1 (where the M queries are defined)
            # The Mashup binary package contains a zip file itself!
            # It starts with 'PK\x03\x04' inside the DataMashup.
            # Let's see if we can find 'PK\x03\x04' in raw_data and extract it!
            pk_index = raw_data.find(b'PK\x03\x04')
            if pk_index != -1:
                log(f"Found ZIP header (PK\\x03\\x04) in DataMashup at index {pk_index}!")
                # Extract the zip file from here
                # A ZIP file goes until the end of the mashup data or we can just try to write it and open it.
                zip_data = raw_data[pk_index:]
                zip_temp = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\mashup_package.zip"
                with open(zip_temp, "wb") as fz:
                    fz.write(zip_data)
                log(f"Wrote mashup package ZIP to {zip_temp}")
                
                # Now try to open this zip and read the queries inside it!
                try:
                    with zipfile.ZipFile(zip_temp, 'r') as mz:
                        log("Contents of Mashup ZIP:")
                        for mname in mz.namelist():
                            log(f"  - {mname} ({len(mz.read(mname))} bytes)")
                            # Formulas are usually in Formulas/Section1.section
                            if "Section1.section" in mname:
                                m_code_raw = mz.read(mname)
                                m_code = decode_data(m_code_raw)
                                log("\n--- POWER QUERY M CODE FROM SECTION1.SECTION ---")
                                log(m_code[:4000]) # log first 4000 chars
                                out.append("\n=== POWER QUERY M CODE FROM SECTION1.SECTION ===")
                                out.append(m_code)
                except Exception as e:
                    log(f"Error opening Mashup ZIP: {e}")
            else:
                log("No ZIP header found in item6.xml binary data.")

    with open(log_path, "w", encoding="utf-8") as lf:
        lf.write("\n".join(out))
    print(f"\nSaved Power Query extract to {log_path}")

if __name__ == "__main__":
    main()
