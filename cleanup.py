import os

files_to_delete = [
    "analyze_excel.py", "check_end_xml.py", "check_string.py", 
    "find_break.py", "find_real_range.py", "fix_excel.py", 
    "fix_excel_clean.py", "hex_analysis.py", "peek_rows.py", 
    "prepare_fix.py", "sample_excel_content.py", "verify_clean.py", 
    "verify_fix.py", "Anexo 1Q Abril 2026_FIXED.xlsx"
]

for f in files_to_delete:
    if os.path.exists(f):
        try:
            os.remove(f)
        except:
            pass
