import zipfile

with zipfile.ZipFile('Anexo 1Q Abril 2026.xlsx', 'r') as zin:
    try:
        content = zin.read('xl/tables/table1.xml')
        print(f"Size: {len(content)}")
        print(repr(content[:500]))
    except Exception as e:
        print(f"Error: {e}")
