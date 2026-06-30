import zipfile
import os
import xml.etree.ElementTree as ET
import re

file_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\Gestion de  Nomina-PortalSolid_V2.xlsm"
log_path = r"C:\Users\amejoramiento6\Desktop\UTILIDADES\scratch\power_queries_output.txt"

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
        custom_xml_files = [name for name in namelist if name.startswith("customXml/item") and name.endswith(".xml")]
        
        log(f"Found {len(custom_xml_files)} custom XML files.")
        
        for cxf in custom_xml_files:
            try:
                content = z.read(cxf).decode("utf-8", errors="ignore")
            except Exception as e:
                log(f"Error reading {cxf}: {e}")
                continue
                
            if "Section1" in content or "shared" in content or "Expression" in content or "Formula" in content:
                log(f"\n==================================================")
                log(f"POWER QUERY DATA FROM {cxf} ({len(content)} bytes)")
                log(f"==================================================")
                
                # Try to extract the M formulas from the DataMashup structure
                # The DataMashup package is actually a binary structure inside the XML, but it often has clear text M code in it.
                # Let's extract portions of code starting with "shared " or "let "
                # We can search for M code using regex
                m_queries = re.findall(r'(\bshared\s+[A-Za-z0-9_ -]+\s*=\s*let\b.*?);', content, re.DOTALL)
                if m_queries:
                    log(f"Found {len(m_queries)} shared queries using basic regex:")
                    for idx, q in enumerate(m_queries):
                        # decode entities
                        q_dec = q.replace('&quot;', '"').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('#(lf)', '\n')
                        log(f"\nQuery #{idx+1}:\n{q_dec}\n" + "-"*50)
                else:
                    # Let's search for "let " and "in " blocks
                    let_blocks = re.findall(r'(let\s+.*?in\s+[A-Za-z0-9_ -]+)', content, re.DOTALL)
                    if let_blocks:
                        log(f"Found {len(let_blocks)} let blocks:")
                        for idx, q in enumerate(let_blocks[:20]): # print first 20
                            q_dec = q.replace('&quot;', '"').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('#(lf)', '\n').replace('#(tab)', '\t')
                            # clean up binary noise around it
                            q_dec_clean = "".join([c if (32 <= ord(c) < 127 or c in '\n\r\t') else ' ' for c in q_dec])
                            # collapse multiple spaces
                            q_dec_clean = re.sub(r' +', ' ', q_dec_clean)
                            log(f"\nBlock #{idx+1}:\n{q_dec_clean[:2000]}...")
                            log("-"*50)
                    else:
                        # Print strings that look like Power Query names or steps
                        log("No let blocks found, displaying raw preview of first 3000 chars:")
                        preview = content[:3000].replace('&quot;', '"').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
                        preview_clean = "".join([c if (32 <= ord(c) < 127 or c in '\n\r\t') else ' ' for c in preview])
                        log(preview_clean)

    with open(log_path, "w", encoding="utf-8") as lf:
        lf.write("\n".join(out))
    print(f"\nSaved Power Query extract to {log_path}")

if __name__ == "__main__":
    main()
