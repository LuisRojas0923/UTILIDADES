
import polars as pl
import os
import re

# ==============================================================================
# SCRIPT DE LECTURA Y COMPARACIÓN
# ==============================================================================
# Este script usa la misma lógica de lectura robusta para cargar el Excel
# y prepararlo para hacer comparaciones (ej. buscar diferencias, cambios, etc.)

FILE_PATH = r"R:\Costos\INFORME DE ORDENES\BASE DE DATOS GENERAL.xlsm"
SHEET_NAME = "MOVIMIENTO_DE ORDENES__2"

def obtener_dataframe_limpio():
    """
    Lee el archivo Excel, encuentra la cabecera automáticamente, normaliza columnas
    y devuelve un DataFrame de Polars listo para comparar.
    """
    print(f"--- 🔍 Leyendo datos para Comparación ---")
    
    # 1. Lectura RAW
    try:
        df = pl.read_excel(FILE_PATH, sheet_name=SHEET_NAME, infer_schema_length=0)
    except Exception as e:
        print(f"❌ Error leyendo archivo: {e}")
        return None

    # 2. Búsqueda Automática de Cabeceras
    keywords = ["ORDEN", "VR. CONTRATADO", "DESCRIPCION", "OP"]
    found_offset = -1
    for i in range(min(20, df.height)):
        row_values = [str(v).upper() for v in df.row(i)]
        matches = sum(1 for kw in keywords if any(kw in val for val in row_values))
        if matches >= 2:
            found_offset = i
            break
    
    if found_offset != -1:
        raw_headers = df.row(found_offset)
        final_headers = []
        seen_headers = {}
        for h in raw_headers:
            h_str = str(h).strip() if h is not None else "col"
            if h_str in ["", "None", "nan", "null"]: h_str = "col"
            h_str = re.sub(r'[^a-zA-Z0-9]', '_', h_str).strip('_')
            if not h_str: h_str = "col"
            original_h = h_str; counter = 1
            while h_str in seen_headers:
                h_str = f"{original_h}_{counter}"; counter += 1
            seen_headers[h_str] = True
            final_headers.append(h_str)
        
        df = df.slice(found_offset + 1)
        df.columns = final_headers
    else:
        print("⚠️ No se encontraron cabeceras claras.")

    # 3. Normalización (Snake Case)
    new_columns = []
    for col in df.columns:
        clean_col = re.sub(r'[^a-z0-9]+', '_', col.strip().lower()).strip('_')
        new_columns.append(clean_col)
    df.columns = new_columns

    # 4. Whitelist (Columnas deseadas)
    wanted_columns = [
        'op', 'orden', 'orden_vieja_nueva', 'clasificacion_de_la_orden', 'cc', 'scc', 'b', 
        'especialidad', 'cod_uen', 'uen', 'sub_indice', 'estado', 'apertura_siigo', 
        'cierre_siigo', 'descripcion_producto_terminado', 'codigo_pp', 'codigo_pt', 
        'categoria_sub_indice', 'cliente', 'ubicaci_n', 'estado_base_contrato', 
        'apertura_base_contratos', 'cierre_base_contratos', 'ingeniero', 'ot_planta', 
        'nit', 'uen_fact', 'vr_contratado', 'descripcion', 'cajas_menores'
    ]
    final_cols = [c for c in df.columns if c in wanted_columns]
    if final_cols:
        df = df.select(final_cols)
    else:
        print("❌ Error: Columnas no coinciden.")
        return None

    # 5. Limpieza Moneda
    for col_name in ["vr_contratado", "cajas_menores"]:
        if col_name in df.columns:
            df = df.with_columns(
                pl.col(col_name).str.replace_all(r"[$. ]", "").str.replace(",", ".").cast(pl.Float64, strict=False)
            )

    print(f"✅ DataFrame listo para comparar. Filas: {df.height}")
    return df

def comparar_datos():
    df_nuevo = obtener_dataframe_limpio()
    if df_nuevo is None: return

    # --- EJEMPLO DE LÓGICA DE COMPARACIÓN ---
    # Aquí puedes añadir lógica para:
    # 1. Leer la tabla actual de Postgres a otro DF (df_viejo)
    # 2. Comparar diferencias (anti-join)
    # 3. Ver registros nuevos o eliminados
    
    print("\n--- Estadísticas del Archivo Leído ---")
    print(df_nuevo.describe())

    # Ejemplo: Contar cuántos por 'estado'
    if 'estado' in df_nuevo.columns:
        print("\n--- Conteo por Estado ---")
        print(df_nuevo.group_by("estado").len().sort("len", descending=True))

if __name__ == "__main__":
    comparar_datos()
