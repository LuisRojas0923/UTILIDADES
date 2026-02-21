
import pandas as pd
from sqlalchemy import create_engine
import time
import os
import re

# ==============================================================================
# CONFIGURACIÓN (PANDAS VERSION)
# ==============================================================================
DB_URI = "postgresql+psycopg2://horarios_user:tu_password_aqui@127.0.0.1:5433/horarios_db"
# Nota: Pandas usa psycopg2 estándar con SQLAlchemy, no ADBC.

FILE_PATH = r"R:\Costos\INFORME DE ORDENES\BASE DE DATOS GENERAL(2).xlsm"
SHEET_NAME = "MOVIMIENTO_DE ORDENES__2"
TABLE_NAME_PG = "basegeneralcostos"

def upload_buffer_pandas():
    print(f"--- Iniciando Carga Buffer (PANDAS) a {TABLE_NAME_PG} ---")
    start_time = time.time()

    # 1. CONEXIÓN
    try:
        engine = create_engine(DB_URI)
        print("✅ Conexión establecida.")
    except Exception as e:
        print(f"❌ Error de conexión: {e}")
        return

    # 2. LECTURA Y BÚSQUEDA DE CABECERAS
    print(f"📂 Leyendo archivo Excel con Pandas: {FILE_PATH}")
    try:
        t0 = time.time()
        
        # Leemos solo las primeras 20 filas para buscar la cabecera
        df_preview = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, nrows=20, header=None)
        
        keywords = ["ORDEN", "VR. CONTRATADO", "DESCRIPCION", "OP"]
        header_row_idx = -1
        
        for i, row in df_preview.iterrows():
            row_str = row.astype(str).str.upper().tolist()
            matches = sum(1 for kw in keywords if any(kw in str(val) for val in row_str))
            
            if matches >= 2:
                header_row_idx = i
                print(f"✅ Cabeceras encontradas en fila (index): {header_row_idx}")
                break
        
        if header_row_idx == -1:
            print("⚠️ No hay cabeceras claras, usando fila 0.")
            header_row_idx = 0

        # Ahora leemos el archivo completo saltando hasta la cabecera
        # Pandas lee mucho más lento que Polars, paciencia.
        df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, header=header_row_idx)
        print(f"✅ Archivo leído en {time.time() - t0:.2f}s. Filas: {len(df)}")

    except Exception as e:
        print(f"❌ Error leyendo archivo: {e}")
        return

    # 3. LIMPIEZA
    print("🧹 Normalizando columnas...")
    try:
        new_columns = []
        for col in df.columns:
            clean = str(col).strip().lower()
            clean = re.sub(r'[^a-z0-9]+', '_', clean).strip('_')
            # Deduplicar
            if clean in new_columns:
                counter = 1
                while f"{clean}_{counter}" in new_columns:
                    counter += 1
                clean = f"{clean}_{counter}"
            new_columns.append(clean)
        
        df.columns = new_columns

        # Whitelist
        wanted_columns = [
            'op', 'orden', 'orden_vieja_nueva', 'clasificacion_de_la_orden', 'cc', 'scc', 'b', 
            'especialidad', 'cod_uen', 'uen', 'sub_indice', 'estado', 'apertura_siigo', 
            'cierre_siigo', 'descripcion_producto_terminado', 'codigo_pp', 'codigo_pt', 
            'categoria_sub_indice', 'cliente', 'ubicaci_n', 'estado_base_contrato', 
            'apertura_base_contratos', 'cierre_base_contratos', 'ingeniero', 'ot_planta', 
            'nit', 'uen_fact', 'vr_contratado', 'descripcion', 'cajas_menores'
        ]
        
        # Filtrar columnas existentes
        existing_cols = [c for c in wanted_columns if c in df.columns]
        df = df[existing_cols]
        
        # Limpieza Moneda
        for col in ["vr_contratado", "cajas_menores"]:
            if col in df.columns:
                # Pandas usa .str vectorizado
                df[col] = df[col].astype(str).str.replace(r'[$. ]', '', regex=True).str.replace(',', '.')
                df[col] = pd.to_numeric(df[col], errors='coerce')

    except Exception as e:
        print(f"❌ Error limpiando datos: {e}")
        return

    # 4. SUBIDA
    print("🚀 Subiendo a PostgreSQL con Pandas (to_sql)...")
    try:
        t1 = time.time()
        # chunksize ayuda memoria, pero inserts siguen siendo 'simples' o multi-value
        df.to_sql(
            TABLE_NAME_PG, 
            engine, 
            if_exists='replace', 
            index=False, 
            chunksize=5000, 
            method='multi' # 'multi' hace INSERT (a,b), (c,d)... es mas rapido que default
        )
        elapsed = time.time() - start_time
        print(f"✅ Carga completada.")
        print(f"⏱️ Tiempo total Pandas: {elapsed:.2f} segundos.")

    except Exception as e:
        print(f"❌ Error subiendo: {e}")

if __name__ == "__main__":
    upload_buffer_pandas()
