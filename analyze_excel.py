import zipfile
import os

file_path = r"c:\Users\amejoramiento6\Desktop\UTILIDADES\Anexo 1Q Abril 2026.xlsx"

try:
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        info_list = zip_ref.infolist()
        # Sort by size (uncompressed)
        info_list.sort(key=lambda x: x.file_size, reverse=True)
        
        print(f"{'File Name':<50} | {'Uncompressed Size':>20} | {'Compressed Size':>20}")
        print("-" * 95)
        for info in info_list[:20]: # Top 20 largest internal files
            print(f"{info.filename:<50} | {info.file_size:>20,} | {info.compress_size:>20,}")
            
except Exception as e:
    print(f"Error: {e}")
