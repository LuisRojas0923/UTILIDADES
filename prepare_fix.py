import zipfile
import shutil
import os
import re

original_file = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"
fixed_file = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026_FIXED.xlsx"
temp_dir = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\temp_excel"

if os.path.exists(temp_dir):
    shutil.rmtree(temp_dir)
os.makedirs(temp_dir)

def fix_excel():
    print("Starting fix process...")
    try:
        with zipfile.ZipFile(original_file, 'r') as zin:
            with zipfile.ZipFile(fixed_file, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == 'xl/worksheets/sheet1.xml':
                        print("Processing sheet1.xml (this may take a minute)...")
                        with zin.open(item.filename) as f_in:
                            # We'll read and write chunks
                            # 1. Update dimension tag
                            # 2. Stop after row 98
                            
                            content_buffer = b""
                            found_sheet_data = False
                            rows_written = 0
                            limit = 98
                            
                            # Read and write bit by bit
                            chunk_size = 1024 * 1024 # 1MB
                            
                            # Phase 1: Header + Dimension
                            header_chunk = f_in.read(chunk_size)
                            # Change dimension ref="A1:CY1048576" to ref="A1:CY98"
                            # We'll use a regex on the decoded string for the first chunk
                            header_str = header_chunk.decode('utf-8', errors='ignore')
                            header_str = re.sub(r'dimension ref="([^"]+):1048576"', rf'dimension ref="\1:{limit}"', header_str)
                            
                            # Find <sheetData>
                            parts = header_str.split('<sheetData>')
                            zout.writestr(item.filename, parts[0] + '<sheetData>')
                            
                            # Now process rows inside <sheetData>
                            # The first chunk might already contain some rows
                            if len(parts) > 1:
                                remaining = parts[1]
                            else:
                                remaining = ""
                            
                            # Function to process rows in a string
                            def process_chunk(s, rows_count):
                                row_matches = list(re.finditer(r'<row r="(\d+)"', s))
                                last_pos = 0
                                for m in row_matches:
                                    r_idx = int(m.group(1))
                                    if r_idx > limit:
                                        # Found a row beyond the limit, stop here
                                        return s[:last_pos], True, rows_count
                                    
                                    # Find the end of this row </row>
                                    end_pos = s.find('</row>', m.end())
                                    if end_pos == -1:
                                        # Incomplete row in this chunk
                                        return s[:last_pos], False, rows_count
                                    
                                    last_pos = end_pos + 6
                                    rows_count += 1
                                
                                return s[:last_pos], False, rows_count

                            # Process the remaining part of the first chunk
                            chunk_to_write, stop, rows_written = process_chunk(remaining, rows_written)
                            zout.open(item.filename, 'a').write(chunk_to_write.encode('utf-8')) # This doesn't work with writestr/open easily
                            # Actually, zipfile.open in 'a' mode is not supported for individual files.
                            # We need to accumulate or use a temp file.
                            
                    else:
                        # Copy other files as is
                        zout.writestr(item, zin.read(item.filename))
    except Exception as e:
        print(f"Error: {e}")

# Re-writing the fix script to be more robust with file streaming
fix_script_v2 = """
import zipfile
import os
import re

original_file = r"c:\\Users\\amejoramiento6\\Desktop\\UTILIDADES\\Anexo 1Q Abril 2026.xlsx"
fixed_file = r"c:\\Users\\amejoramiento6\\Desktop\\UTILIDADES\\Anexo 1Q Abril 2026_FIXED.xlsx"

def fix():
    with zipfile.ZipFile(original_file, 'r') as zin:
        with zipfile.ZipFile(fixed_file, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'xl/worksheets/sheet1.xml':
                    print("Cleaning sheet1.xml...")
                    with zin.open(item.filename) as f_in, zout.open(item.filename, 'w') as f_out:
                        limit = 98
                        found_limit = False
                        
                        # Read the header and fix dimension
                        # We'll read until <sheetData>
                        buffer = b""
                        while b'<sheetData>' not in buffer:
                            chunk = f_in.read(4096)
                            if not chunk: break
                            buffer += chunk
                        
                        head, tail = buffer.split(b'<sheetData>', 1)
                        head_str = head.decode('utf-8', errors='ignore')
                        head_str = re.sub(r'dimension ref="([^"]+):1048576"', rf'dimension ref="\\1:{limit}"', head_str)
                        f_out.write(head_str.encode('utf-8') + b'<sheetData>')
                        
                        # Process rows
                        buffer = tail
                        row_pattern = re.compile(br'<row r="(\\d+)"')
                        
                        while True:
                            # Process what's in buffer
                            while True:
                                # Find a row
                                match = row_pattern.search(buffer)
                                if not match: break
                                
                                r_idx = int(match.group(1))
                                if r_idx > limit:
                                    found_limit = True
                                    break
                                
                                # Find end of row
                                end_tag = b'</row>'
                                end_pos = buffer.find(end_tag, match.end())
                                if end_pos == -1: break # Need more data
                                
                                # Write this row
                                f_out.write(buffer[:end_pos + len(end_tag)])
                                buffer = buffer[end_pos + len(end_tag):]
                            
                            if found_limit or not buffer:
                                break
                                
                            new_chunk = f_in.read(1024*1024)
                            if not new_chunk: break
                            buffer += new_chunk
                        
                        # Close tags
                        f_out.write(b'</sheetData></worksheet>')
                
                elif item.filename == 'xl/calcChain.xml':
                    # This file is also huge and probably contains invalid refs now
                    # Excel will regenerate it if missing. Let's just skip it to save space.
                    print("Skipping calcChain.xml (Excel will regenerate it)...")
                    continue
                else:
                    zout.writestr(item, zin.read(item.filename))
    print(f"Done! Fixed file created at: {fixed_file}")
    print(f"New size: {os.path.getsize(fixed_file) / 1024:.2f} KB")

fix()
"""

with open(r"c:\Users\amejoramiento6\Desktop\UTILIDADES\fix_excel_final.py", "w") as f:
    f.write(fix_script_v2)
