
import polars as pl
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from sqlalchemy import create_engine


# ==============================================================================
# CONFIGURACION DEL SCRIPT
# ==============================================================================

# Conexion a PostgreSQL
DB_URI = "postgresql://postgres:AdminSolid2025@192.168.0.21:5432/solid"

# Tabla destino en PostgreSQL
TABLE_NAME_PG = "basegeneralcostos"

# ------------------------------------------------------------------------------
# ARCHIVO PRIMARIO (Base General de Costos)
# ------------------------------------------------------------------------------
FILE_PRIMARY = r"\\192.168.0.3\Procesos Comunes SGI\Costos\INFORME DE ORDENES\BASE DE DATOS GENERAL.xlsm"
SHEET_PRIMARY = "MOVIMIENTO_DE ORDENES__2"

# ------------------------------------------------------------------------------
# ARCHIVO SECUNDARIO (Informe Gestion Postventa)
# ------------------------------------------------------------------------------
FILE_SECONDARY = r"\\192.168.0.3\Postventa\MANTENIMIENTO Y SERVICIO POSTVENTA\- GESTION ORDENES DE SERVICIO\CENTRO LOGÍSTICO\MTZ-SPT-02 Informe Gestion Postventa V1.xlsm"
SHEET_SECONDARY = "Info_GesPostV"
HEADER_ROW_SECONDARY = 3  # Fila donde estan los encabezados (1-indexed, fila 3 en Excel)

# ------------------------------------------------------------------------------
# MAPEO DE COLUMNAS: Secundario -> Primario
# ------------------------------------------------------------------------------
# Las claves son los nombres normalizados del secundario
# Los valores son los nombres del esquema primario
# NOTA: Solo mapeamos columnas de las primeras 16 posiciones (0-15) para optimizar lectura
COLUMN_MAPPING_SECONDARY = {
    'numero_os': 'orden',                      # Posicion 2
    'clasificacion_servicio': 'clasificacion_de_la_orden',  # Posicion 7
    'cc': 'cc',                                # Posicion 8
    'subc_costos': 'scc',                      # Posicion 9
    'subindice': 'sub_indice',                 # Posicion 10
    'fecha_apertura': 'apertura_siigo',        # Posicion 12
    'fecha_de_cierre': 'cierre_siigo',         # Posicion 13
    'cliente': 'cliente',                      # Posicion 4
    'ciudad': 'ubicaci_n',                     # Posicion 5
    'valor_presupuestado': 'vr_contratado',    # Posicion 15
}

# Columnas a leer del secundario (solo las primeras 16 para optimizar)
SECONDARY_COLUMNS_TO_READ = 16

# Columnas del esquema primario que no existen en secundario (seran NULL)
# Incluye 'estado' y 'descripcion' que estan en posiciones lejanas (36, 38)
NULL_COLUMNS_FOR_SECONDARY = [
    'op', 'orden_vieja_nueva', 'especialidad', 'cod_uen',
    'descripcion_producto_terminado', 'codigo_pp', 'codigo_pt',
    'categoria_sub_indice', 'estado_base_contrato', 'apertura_base_contratos',
    'cierre_base_contratos', 'ingeniero', 'ot_planta', 'nit', 'uen_fact', 'cajas_menores',
    'estado', 'descripcion'  # Omitidas del secundario para optimizar lectura
]

# Esquema completo (30 columnas originales + 1 fuente)
FULL_SCHEMA = [
    'op', 'orden', 'orden_vieja_nueva', 'clasificacion_de_la_orden', 'cc', 'scc', 'b', 
    'especialidad', 'cod_uen', 'uen', 'sub_indice', 'estado', 'apertura_siigo', 
    'cierre_siigo', 'descripcion_producto_terminado', 'codigo_pp', 'codigo_pt', 
    'categoria_sub_indice', 'cliente', 'ubicaci_n', 'estado_base_contrato', 
    'apertura_base_contratos', 'cierre_base_contratos', 'ingeniero', 'ot_planta', 
    'nit', 'uen_fact', 'vr_contratado', 'descripcion', 'cajas_menores', 'fuente'
]


# ==============================================================================
# FUNCIONES DE UTILIDAD
# ==============================================================================

def normalize_column_name(col: str) -> str:
    """Normaliza un nombre de columna a snake_case"""
    clean_col = col.strip().lower()
    clean_col = re.sub(r'[^a-z0-9]+', '_', clean_col)
    clean_col = clean_col.strip('_')
    return clean_col


def find_header_row(df: pl.DataFrame, keywords: list, max_rows: int = 20) -> int:
    """Busca la fila que contiene los encabezados basandose en palabras clave"""
    for i in range(min(max_rows, df.height)):
        row_values = [str(v).upper() for v in df.row(i)]
        matches = sum(1 for kw in keywords if any(kw in val for val in row_values))
        if matches >= 2:
            return i
    return -1


def clean_and_deduplicate_headers(raw_headers: tuple) -> list:
    """Limpia y deduplica nombres de columnas"""
    final_headers = []
    seen_headers = {}
    
    for h in raw_headers:
        h_str = str(h).strip() if h is not None else "col"
        if h_str in ["", "None", "nan", "null"]:
            h_str = "col"
        
        h_str = re.sub(r'[^a-zA-Z0-9]', '_', h_str).strip('_')
        if not h_str: 
            h_str = "col"

        original_h = h_str
        counter = 1
        while h_str in seen_headers:
            h_str = f"{original_h}_{counter}"
            counter += 1
        
        seen_headers[h_str] = True
        final_headers.append(h_str)
    
    return final_headers


# ==============================================================================
# LECTURA DEL ARCHIVO PRIMARIO
# ==============================================================================

def read_primary_excel() -> pl.DataFrame:
    """Lee el archivo primario (Base General de Costos) y agrega columna fuente"""
    print(f"\n{'='*60}")
    print(f"[PRIMARIO] Leyendo: {FILE_PRIMARY}")
    print(f"           Hoja: {SHEET_PRIMARY}")
    print(f"{'='*60}")
    
    t0 = time.time()
    
    # Leer Excel
    df = pl.read_excel(
        FILE_PRIMARY, 
        sheet_name=SHEET_PRIMARY,
        infer_schema_length=0
    )
    print(f"  Archivo leido bruto. Filas: {df.height}")
    
    # Buscar cabeceras
    keywords = ["ORDEN", "VR. CONTRATADO", "DESCRIPCION", "OP"]
    found_offset = find_header_row(df, keywords)
    
    if found_offset != -1:
        raw_headers = df.row(found_offset)
        final_headers = clean_and_deduplicate_headers(raw_headers)
        df = df.slice(found_offset + 1)
        df.columns = final_headers
        print(f"  Cabeceras encontradas en indice {found_offset}")
    else:
        print("  ADVERTENCIA: No se encontraron cabeceras esperadas")
    
    # Normalizar nombres de columnas
    df.columns = [normalize_column_name(c) for c in df.columns]
    
    # Filtrar solo columnas deseadas (sin 'fuente' por ahora)
    wanted_columns = [c for c in FULL_SCHEMA if c != 'fuente']
    final_cols = [c for c in df.columns if c in wanted_columns]
    
    missing = set(wanted_columns) - set(final_cols)
    if missing:
        print(f"  Alerta: Faltan columnas: {missing}")
    
    df = df.select(final_cols)
    
    # Limpiar columnas de moneda
    cols_to_clean = ["vr_contratado", "cajas_menores"]
    for col_name in cols_to_clean:
        if col_name in df.columns:
            df = df.with_columns(
                pl.col(col_name)
                .str.replace_all(r"[$. ]", "") 
                .str.replace(",", ".")          
                .cast(pl.Float64, strict=False) 
            )
    
    # Agregar columna fuente
    df = df.with_columns(pl.lit("BASE_GENERAL").alias("fuente"))
    
    print(f"  Lectura completada en {time.time() - t0:.2f}s")
    print(f"  Filas: {df.height}, Columnas: {df.width}")
    
    return df


# ==============================================================================
# LECTURA DEL ARCHIVO SECUNDARIO
# ==============================================================================

def read_secondary_excel() -> pl.DataFrame:
    """Lee el archivo secundario (Informe Gestion Postventa) - OPTIMIZADO"""
    print(f"\n{'='*60}")
    print(f"[SECUNDARIO] Leyendo: {FILE_SECONDARY}")
    print(f"             Hoja: {SHEET_SECONDARY}")
    print(f"             Solo primeras {SECONDARY_COLUMNS_TO_READ} columnas (optimizado)")
    print(f"{'='*60}")
    
    t0 = time.time()
    
    # Leer Excel
    df = pl.read_excel(
        FILE_SECONDARY, 
        sheet_name=SHEET_SECONDARY,
        infer_schema_length=0
    )
    print(f"  Archivo leido bruto. Filas: {df.height}, Cols: {df.width}")
    
    # OPTIMIZACION: Seleccionar solo las primeras N columnas ANTES de procesar
    df = df.select(df.columns[:SECONDARY_COLUMNS_TO_READ])
    print(f"  Columnas recortadas a: {df.width}")
    
    # Usar fila especifica para encabezados (HEADER_ROW_SECONDARY es 1-indexed)
    header_idx = HEADER_ROW_SECONDARY - 1  # Convertir a 0-indexed
    
    if header_idx < df.height:
        raw_headers = df.row(header_idx)
        final_headers = clean_and_deduplicate_headers(raw_headers)
        df = df.slice(header_idx + 1)
        df.columns = final_headers
        print(f"  Cabeceras tomadas de fila {HEADER_ROW_SECONDARY}")
    else:
        print(f"  ERROR: Fila {HEADER_ROW_SECONDARY} no existe en el archivo")
        return pl.DataFrame()
    
    # Normalizar nombres de columnas
    df.columns = [normalize_column_name(c) for c in df.columns]
    
    print(f"  Columnas: {df.columns}")
    print(f"  Lectura completada en {time.time() - t0:.2f}s")
    print(f"  Filas: {df.height}")
    
    return df


# ==============================================================================
# MAPEO DE COLUMNAS SECUNDARIO -> ESQUEMA PRIMARIO
# ==============================================================================

def map_secondary_to_schema(df_secondary: pl.DataFrame) -> pl.DataFrame:
    """Mapea las columnas del secundario al esquema del primario"""
    print(f"\n[MAPEO] Transformando columnas del secundario al esquema primario...")
    
    # Crear DataFrame con columnas mapeadas
    mapped_cols = []
    
    for sec_col, pri_col in COLUMN_MAPPING_SECONDARY.items():
        if sec_col in df_secondary.columns:
            mapped_cols.append(pl.col(sec_col).alias(pri_col))
            print(f"  {sec_col} -> {pri_col}")
        else:
            # Si no existe, crear columna NULL
            mapped_cols.append(pl.lit(None).alias(pri_col))
            print(f"  {sec_col} -> {pri_col} (NULL - no encontrada)")
    
    # Seleccionar solo las columnas mapeadas
    df_mapped = df_secondary.select(mapped_cols)
    
    # Agregar columnas NULL para las que no tienen equivalente
    for null_col in NULL_COLUMNS_FOR_SECONDARY:
        df_mapped = df_mapped.with_columns(pl.lit(None).alias(null_col))
        print(f"  (NULL) -> {null_col}")
    
    # Logica condicional para columna 'b' (basada en 'scc')
    # "B", each if 'scc' = 20 then "31" else if scc = 10 then "30" else null
    df_mapped = df_mapped.with_columns(
        pl.when(pl.col("scc").cast(pl.Int64, strict=False) == 20).then(pl.lit("31"))
        .when(pl.col("scc").cast(pl.Int64, strict=False) == 10).then(pl.lit("30"))
        .otherwise(None)
        .alias("b")
    )
    print("  (IF-LOGIC) -> b")

    # Logica condicional para columna 'uen' (basada en 'orden')
    # "uen", each if orden <= 17000 then "ADN" else null
    df_mapped = df_mapped.with_columns(
        pl.when(pl.col("orden").cast(pl.Int64, strict=False) <= 17000).then(pl.lit("ADN"))
        .otherwise(None)
        .alias("uen")
    )
    print("  (IF-LOGIC) -> uen")

    # Limpiar columna de moneda vr_contratado
    if "vr_contratado" in df_mapped.columns:
        df_mapped = df_mapped.with_columns(
            pl.col("vr_contratado")
            .cast(pl.Utf8)
            .str.replace_all(r"[$. ]", "") 
            .str.replace(",", ".")          
            .cast(pl.Float64, strict=False) 
        )
    
    # Agregar columna fuente
    df_mapped = df_mapped.with_columns(pl.lit("POSTVENTA").alias("fuente"))
    
    # Reordenar columnas segun esquema completo
    final_cols = [c for c in FULL_SCHEMA if c in df_mapped.columns]
    df_mapped = df_mapped.select(final_cols)
    
    print(f"  Columnas finales: {len(df_mapped.columns)}")
    
    return df_mapped


# ==============================================================================
# MERGE DE DATAFRAMES
# ==============================================================================

def merge_dataframes(df_primary: pl.DataFrame, df_secondary_mapped: pl.DataFrame) -> pl.DataFrame:
    """Filtra ordenes nuevas del secundario y las concatena con el primario"""
    print(f"\n[MERGE] Combinando DataFrames...")
    
    # Obtener ordenes existentes en el primario
    ordenes_primario = set(df_primary.select("orden").to_series().to_list())
    print(f"  Ordenes en primario: {len(ordenes_primario)}")
    
    # Filtrar secundario: solo ordenes que NO estan en primario
    df_secondary_filtered = df_secondary_mapped.filter(
        ~pl.col("orden").is_in(list(ordenes_primario))
    )
    print(f"  Ordenes nuevas en secundario: {df_secondary_filtered.height}")
    
    if df_secondary_filtered.height == 0:
        print("  No hay ordenes nuevas para agregar")
        return df_primary
    
    # Asegurar que ambos DataFrames tengan las mismas columnas en el mismo orden
    common_cols = [c for c in df_primary.columns if c in df_secondary_filtered.columns]
    
    df_primary_aligned = df_primary.select(common_cols)
    df_secondary_aligned = df_secondary_filtered.select(common_cols)
    
    # Concatenar
    df_merged = pl.concat([df_primary_aligned, df_secondary_aligned])
    
    print(f"  Total filas despues del merge: {df_merged.height}")
    print(f"    - Del primario (BASE_GENERAL): {df_primary.height}")
    print(f"    - Del secundario (POSTVENTA): {df_secondary_filtered.height}")
    
    return df_merged


# ==============================================================================
# SUBIDA A POSTGRESQL
# ==============================================================================

def upload_to_postgres(df: pl.DataFrame):
    """Sube el DataFrame a PostgreSQL"""
    print(f"\n[UPLOAD] Subiendo datos a PostgreSQL...")
    print(f"  Tabla destino: {TABLE_NAME_PG}")
    print(f"  Filas a insertar: {df.height}")
    print(f"  Columnas: {df.width}")
    
    t0 = time.time()
    
    try:
        df.write_database(
            table_name=TABLE_NAME_PG, 
            connection=DB_URI, 
            if_table_exists="replace",
            engine="adbc"
        )
        print(f"  Usando motor ADBC (Ultra Rapido)")
        print(f"  Carga completada en {time.time() - t0:.2f}s")
        
    except Exception as adbc_error:
        print(f"  Error ADBC: {adbc_error}")
        print("  Verifica que PostgreSQL este activo en el puerto configurado.")
        raise


# ==============================================================================
# LECTURA PARALELA
# ==============================================================================

def read_files_parallel():
    """Lee y procesa archivos en paralelo usando un pipeline para el secundario"""
    print(f"\n[PARALELO] Iniciando pipeline de procesamiento...")
    t0 = time.time()
    
    with ThreadPoolExecutor(max_workers=2) as executor:
        # Tarea 1: Leer primario (pesado)
        future_primary = executor.submit(read_primary_excel)
        
        # Tarea 2: Leer Y MAPEAR secundario (ligero, se hace mientras el primario carga)
        def secondary_pipeline():
            df_sec = read_secondary_excel()
            return map_secondary_to_schema(df_sec)
            
        future_secondary_mapped = executor.submit(secondary_pipeline)
        
        # Esperar resultados
        df_primary = future_primary.result()
        df_secondary_mapped = future_secondary_mapped.result()
    
    elapsed = time.time() - t0
    print(f"\n[PARALELO] Pipeline completado en {elapsed:.2f}s")
    
    return df_primary, df_secondary_mapped


# ==============================================================================
# FUNCION PRINCIPAL
# ==============================================================================

def upload_buffer_with_merge():
    """Funcion principal: Pipeline paralelo de lectura/mapeo, merge y upload"""
    print("\n" + "="*70)
    print("  CARGA BUFFER CON PIPELINE PARALELO")
    print("="*70)
    
    start_time = time.time()
    
    try:
        # 1. Ejecutar el pipeline paralelo (Lectura P + Lectura S + Mapeo S)
        df_primary, df_secondary_mapped = read_files_parallel()
        
        # 2. Merge: agregar ordenes nuevas del secundario
        df_merged = merge_dataframes(df_primary, df_secondary_mapped)
        
        # 3. Subir a PostgreSQL
        upload_to_postgres(df_merged)
        
        elapsed = time.time() - start_time
        print(f"\n{'='*70}")
        print(f"  COMPLETADO EXITOSAMENTE")
        print(f"  Tiempo total: {elapsed:.2f} segundos")
        print(f"{'='*70}\n")
        
    except Exception as e:
        print(f"\nERROR FATAL: {e}")
        raise


if __name__ == "__main__":
    import sys
    try:
        upload_buffer_with_merge()
        sys.exit(0)
    except Exception as e:
        print(f"ERROR FATAL: {e}")
        sys.exit(1)
