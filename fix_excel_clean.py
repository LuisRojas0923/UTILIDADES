import zipfile
import os
import re

original_file = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"
fixed_file = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026_CLEAN.xlsx"

def fix():
    if os.path.exists(fixed_file):
        try:
            os.remove(fixed_file)
        except:
            pass # Hope for the best or use a different name if needed

    with zipfile.ZipFile(original_file, 'r') as zin:
        with zipfile.ZipFile(fixed_file, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'xl/worksheets/sheet1.xml':
                    print("Cleaning sheet1.xml...")
                    with zin.open(item.filename) as f_in, zout.open(item.filename, 'w') as f_out:
                        limit = 98
                        found_limit = False
                        
                        buffer = b""
                        while b'<sheetData>' not in buffer:
                            chunk = f_in.read(4096)
                            if not chunk: break
                            buffer += chunk
                        
                        if b'<sheetData>' in buffer:
                            head, tail = buffer.split(b'<sheetData>', 1)
                            head_str = head.decode('utf-8', errors='ignore')
                            head_str = re.sub(r'dimension ref="[A-Z0-9:]+"', f'dimension ref="A1:CY{limit}"', head_str)
                            f_out.write(head_str.encode('utf-8') + b'<sheetData>')
                            buffer = tail
                        
                        row_pattern = re.compile(br'<row r="(\d+)"')
                        while True:
                            while True:
                                match = row_pattern.search(buffer)
                                if not match: break
                                r_idx = int(match.group(1))
                                if r_idx > limit:
                                    found_limit = True
                                    break
                                end_tag = b'</row>'
                                end_pos = buffer.find(end_tag, match.end())
                                if end_pos == -1: break
                                f_out.write(buffer[:end_pos + len(end_tag)])
                                buffer = buffer[end_pos + len(end_tag):]
                            if found_limit: break
                            new_chunk = f_in.read(1024*1024)
                            if not new_chunk: break
                            buffer += new_chunk
                        
                        while b'</sheetData>' not in buffer:
                            new_chunk = f_in.read(1024*1024)
                            if not new_chunk: break
                            buffer += new_chunk
                        
                        if b'</sheetData>' in buffer:
                            _, rest = buffer.split(b'</sheetData>', 1)
                            f_out.write(b'</sheetData>' + rest)
                            while True:
                                chunk = f_in.read(1024*1024)
                                if not chunk: break
                                f_out.write(chunk)
                elif item.filename == 'xl/calcChain.xml':
                    continue
                else:
                    zout.writestr(item, zin.read(item.filename))
    print(f"Done! New size: {os.path.getsize(fixed_file) / 1024:.2f} KB")

if __name__ == "__main__":
    fix()
