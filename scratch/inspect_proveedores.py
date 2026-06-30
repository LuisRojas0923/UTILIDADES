from sqlalchemy import create_engine, text
import polars as pl
import fastexcel

uri = "postgresql://postgres:AdminSolid2025@192.168.0.21:5432/solid"
e = create_engine(uri)
with e.connect() as c:
    print("=== Tabla proveedor (solid) ===")
    print("count:", c.execute(text("SELECT COUNT(*) FROM proveedor")).scalar())
    print("columns:", [r[0] for r in c.execute(text(
        "SELECT column_name FROM information_schema.columns WHERE table_name='proveedor' ORDER BY ordinal_position"
    )).fetchall()])
    print("sample nit,nombre:")
    for r in c.execute(text("SELECT nit, nombre FROM proveedor ORDER BY codigo LIMIT 8")).fetchall():
        print(" ", r)

path = r"\\192.168.0.3\control presupuestal\CATALOGO DE PRODUCTOS\CATALOGO DE PRODUCTOS.xlsm"
print("\n=== Excel PROVEEDOR PRINC ===")
excel = fastexcel.read_excel(path)
df = excel.load_sheet_by_name("PROVEEDOR PRINC").to_polars()
print("shape:", df.height, df.width)
for i in range(min(12, df.height)):
    row = [str(x) if x is not None else "" for x in df.row(i)]
    print(f"row {i}:", row[:7])
