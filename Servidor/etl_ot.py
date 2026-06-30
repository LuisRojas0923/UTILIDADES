import time

import fastexcel
import polars as pl
from sqlalchemy import create_engine, text

from etl_common import (
    COLUMN_MAPPING_MEMOFICHAS,
    COLUMN_MAPPING_SECONDARY,
    DB_URI,
    FILE_MEMOFICHAS,
    FILE_PRIMARY,
    FILE_SECONDARY,
    FULL_SCHEMA,
    HEADER_ROW_SECONDARY,
    NULL_COLUMNS_FOR_MEMOFICHAS,
    NULL_COLUMNS_FOR_SECONDARY,
    SECONDARY_COLUMNS_TO_READ,
    SHEET_MEMOFICHAS,
    SHEET_PRIMARY,
    SHEET_SECONDARY,
    TABLE_NAME_PG,
    clean_and_deduplicate_headers,
    ensure_utf8_encoding,
    fill_null_values,
    find_header_row,
    log,
    normalize_column_name,
)

# ==============================================================================
# LECTURA DEL ARCHIVO PRIMARIO
# ==============================================================================

def read_primary_excel() -> pl.DataFrame:
    """Lee el archivo primario (Base General de Costos) y agrega columna fuente"""
    log(f"[PRIMARIO] Leyendo: {FILE_PRIMARY}")
    log(f"           Hoja: {SHEET_PRIMARY}")
    
    t0 = time.time()
    
    # Leer Excel optimizado usando fastexcel directamente
    excel = fastexcel.read_excel(FILE_PRIMARY)
    df = excel.load_sheet_by_name(SHEET_PRIMARY).to_polars()
    log(f"  Archivo leido bruto. Filas: {df.height}")
    
    # Buscar cabeceras
    keywords = ["ORDEN", "VR. CONTRATADO", "DESCRIPCION", "OP"]
    found_offset = find_header_row(df, keywords)
    
    if found_offset != -1:
        raw_headers = df.row(found_offset)
        final_headers = clean_and_deduplicate_headers(raw_headers)
        df = df.slice(found_offset + 1)
        df.columns = final_headers
        log(f"  Cabeceras encontradas en indice {found_offset}")
    else:
        log("  ADVERTENCIA: No se encontraron cabeceras esperadas", "warning")
    
    # Normalizar nombres de columnas
    df.columns = [normalize_column_name(c) for c in df.columns]
    
    # Manejar caso especial de 'ubicaci_n' que viene asi en el Excel
    if "ubicaci_n" in df.columns:
        df = df.rename({"ubicaci_n": "ubicacion"})
        log("  Renombrando columna 'ubicaci_n' a 'ubicacion' (mapeo manual)")
    
    # Filtrar solo columnas deseadas (sin 'fuente' por ahora)
    wanted_columns = [c for c in FULL_SCHEMA if c != 'fuente']
    final_cols = [c for c in df.columns if c in wanted_columns]
    
    missing = set(wanted_columns) - set(final_cols)
    if missing:
        log(f"  Alerta: Faltan columnas: {missing}", "warning")
    
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
    
    log(f"  Lectura completada en {time.time() - t0:.2f}s")
    log(f"  Filas: {df.height}, Columnas: {df.width}")
    
    return df


# ==============================================================================
# LECTURA DEL ARCHIVO SECUNDARIO
# ==============================================================================

def read_secondary_excel() -> pl.DataFrame:
    """Lee el archivo secundario (Informe Gestion Postventa) - OPTIMIZADO"""
    log(f"[SECUNDARIO] Leyendo: {FILE_SECONDARY}")
    log(f"             Hoja: {SHEET_SECONDARY}")
    log(f"             Solo primeras {SECONDARY_COLUMNS_TO_READ} columnas (optimizado)")
    
    t0 = time.time()
    
    # Leer Excel optimizado usando fastexcel directamente
    excel = fastexcel.read_excel(FILE_SECONDARY)
    df = excel.load_sheet_by_name(SHEET_SECONDARY).to_polars()
    log(f"  Archivo leido bruto. Filas: {df.height}, Cols: {df.width}")
    
    # OPTIMIZACION: Seleccionar solo las primeras N columnas ANTES de procesar
    df = df.select(df.columns[:SECONDARY_COLUMNS_TO_READ])
    log(f"  Columnas recortadas a: {df.width}")
    
    # Usar fila especifica para encabezados (HEADER_ROW_SECONDARY es 1-indexed)
    header_idx = HEADER_ROW_SECONDARY - 1  # Convertir a 0-indexed
    
    if header_idx < df.height:
        raw_headers = df.row(header_idx)
        final_headers = clean_and_deduplicate_headers(raw_headers)
        df = df.slice(header_idx + 1)
        df.columns = final_headers
        log(f"  Cabeceras tomadas de fila {HEADER_ROW_SECONDARY}")
    else:
        log(f"  ERROR: Fila {HEADER_ROW_SECONDARY} no existe en el archivo", "error")
        return pl.DataFrame()
    
    # Normalizar nombres de columnas
    df.columns = [normalize_column_name(c) for c in df.columns]
    
    log(f"  Columnas: {df.columns}")
    log(f"  Lectura completada en {time.time() - t0:.2f}s")
    log(f"  Filas: {df.height}")
    
    return df


# ==============================================================================
# MAPEO DE COLUMNAS SECUNDARIO -> ESQUEMA PRIMARIO
# ==============================================================================

def map_secondary_to_schema(df_secondary: pl.DataFrame) -> pl.DataFrame:
    """Mapea las columnas del secundario al esquema del primario"""
    log(f"[MAPEO] Transformando columnas del secundario al esquema primario...")
    
    # Crear DataFrame con columnas mapeadas
    mapped_cols = []
    
    for sec_col, pri_col in COLUMN_MAPPING_SECONDARY.items():
        if sec_col in df_secondary.columns:
            mapped_cols.append(pl.col(sec_col).alias(pri_col))
            log(f"  {sec_col} -> {pri_col}")
        else:
            # Si no existe, crear columna NULL
            mapped_cols.append(pl.lit(None).alias(pri_col))
            log(f"  {sec_col} -> {pri_col} (NULL - no encontrada)", "warning")
    
    # Seleccionar solo las columnas mapeadas
    df_mapped = df_secondary.select(mapped_cols)
    
    # Agregar columnas NULL para las que no tienen equivalente
    for null_col in NULL_COLUMNS_FOR_SECONDARY:
        df_mapped = df_mapped.with_columns(pl.lit(None).alias(null_col))
        log(f"  (NULL) -> {null_col}")
    
    # Logica condicional para columna 'b' (basada en 'scc')
    # "B", each if 'scc' = 20 then "31" else if scc = 10 then "30" else null
    df_mapped = df_mapped.with_columns(
        pl.when(pl.col("scc").cast(pl.Int64, strict=False) == 20).then(pl.lit("31"))
        .when(pl.col("scc").cast(pl.Int64, strict=False) == 10).then(pl.lit("30"))
        .otherwise(None)
        .alias("b")
    )
    log("  (IF-LOGIC) -> b")

    # Logica condicional para columna 'uen' (basada en 'orden')
    # "uen", each if orden <= 17000 then "ADN" else null
    df_mapped = df_mapped.with_columns(
        pl.when(pl.col("orden").cast(pl.Int64, strict=False) <= 17000).then(pl.lit("ADN"))
        .otherwise(None)
        .alias("uen")
    )
    log("  (IF-LOGIC) -> uen")

    df_mapped = df_mapped.with_columns(pl.lit("PRODUCE").alias("categoria_sub_indice"))
    log("  (POSTVENTA) -> categoria_sub_indice = PRODUCE")

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
    
    log(f"  Columnas finales: {len(df_mapped.columns)}")
    
    return df_mapped


# ==============================================================================
# LECTURA Y MAPEO MEMOFICHAS (Control Presupuestal)
# ==============================================================================

def read_memofichas_excel() -> pl.DataFrame | None:
    """Lee el archivo MEMOFICHAS (Control Presupuestal). Si no existe o no es accesible, retorna None (MEMOFICHAS opcional)."""
    log(f"[MEMOFICHAS] Leyendo: {FILE_MEMOFICHAS}")
    log(f"             Hoja: {SHEET_MEMOFICHAS}")
    
    t0 = time.time()
    
    try:
        excel = fastexcel.read_excel(FILE_MEMOFICHAS)
        df = excel.load_sheet_by_name(SHEET_MEMOFICHAS).to_polars()
    except Exception as e:
        log(f"  MEMOFICHAS omitido: archivo no encontrado o no accesible. ({e})", "warning")
        log(f"  La sincronizacion continua con Primario + Postventa + Catalogo.", "warning")
        return None
    log(f"  Archivo leido bruto. Filas: {df.height}, Cols: {df.width}")
    
    # Buscar fila de cabeceras por palabras clave (OP, ORDEN, CC, SCC)
    keywords = ["ORDEN", "OP", "CC", "SCC"]
    found_offset = find_header_row(df, keywords)
    
    if found_offset != -1:
        raw_headers = df.row(found_offset)
        final_headers = clean_and_deduplicate_headers(raw_headers)
        df = df.slice(found_offset + 1)
        df.columns = final_headers
        log(f"  Cabeceras encontradas en indice {found_offset}")
    else:
        log("  ADVERTENCIA: No se encontraron cabeceras esperadas (ORDEN, OP, CC). Usando fila 0.", "warning")
        if df.height > 0:
            raw_headers = df.row(0)
            final_headers = clean_and_deduplicate_headers(raw_headers)
            df = df.slice(1)
            df.columns = final_headers
    
    # Normalizar nombres de columnas (snake_case, sin acentos)
    df.columns = [normalize_column_name(c) for c in df.columns]
    
    log(f"  Columnas: {df.columns}")
    log(f"  Lectura completada en {time.time() - t0:.2f}s. Filas: {df.height}")
    
    return df


def map_memofichas_to_schema(df_memofichas: pl.DataFrame) -> pl.DataFrame:
    """Mapea las columnas de MEMOFICHAS al esquema basegeneralcostos (FULL_SCHEMA).
    Regla sub_indice: por niveles de prioridad si están vacíos — SUBINDICE_2 > SUBINDICE_1 > SUBINDICE (primer no vacío).
    """
    log(f"[MAPEO MEMOFICHAS] Transformando al esquema basegeneralcostos...")
    
    def _non_empty(expr):
        return expr.fill_null("").cast(pl.Utf8).str.strip_chars().str.len_chars() > 0
    
    mapped_cols = []
    for mem_col, pri_col in COLUMN_MAPPING_MEMOFICHAS.items():
        # Regla por niveles: sub_indice = primer no vacío (subindice_2 -> subindice_1 -> subindice)
        if pri_col == "sub_indice":
            if "subindice_2" in df_memofichas.columns and "subindice_1" in df_memofichas.columns:
                cascade = (
                    pl.when(_non_empty(pl.col("subindice_2"))).then(pl.col("subindice_2").fill_null(""))
                    .when(_non_empty(pl.col("subindice_1"))).then(pl.col("subindice_1").fill_null(""))
                    .otherwise(pl.col("subindice").fill_null(""))
                )
            elif "subindice_1" in df_memofichas.columns:
                cascade = pl.when(_non_empty(pl.col("subindice_1"))).then(pl.col("subindice_1").fill_null("")).otherwise(pl.col("subindice").fill_null(""))
            else:
                cascade = pl.col("subindice").fill_null("")
            mapped_cols.append(cascade.alias("sub_indice"))
            log("  sub_indice = primer no vacío (subindice_2 -> subindice_1 -> subindice)")
            continue
        if mem_col in df_memofichas.columns:
            mapped_cols.append(pl.col(mem_col).alias(pri_col))
            log(f"  {mem_col} -> {pri_col}")
        else:
            mapped_cols.append(pl.lit(None).alias(pri_col))
            log(f"  {mem_col} -> {pri_col} (NULL - no encontrada)", "warning")
    
    df_mapped = df_memofichas.select(mapped_cols)
    
    for null_col in NULL_COLUMNS_FOR_MEMOFICHAS:
        df_mapped = df_mapped.with_columns(pl.lit(None).alias(null_col))
    
    # Columna fuente
    df_mapped = df_mapped.with_columns(pl.lit("MEMOFICHAS").alias("fuente"))
    
    # Reordenar segun FULL_SCHEMA
    final_cols = [c for c in FULL_SCHEMA if c in df_mapped.columns]
    df_mapped = df_mapped.select(final_cols)
    
    log(f"  Columnas finales: {len(df_mapped.columns)}")
    return df_mapped


# ==============================================================================
# MERGE DE DATAFRAMES
# ==============================================================================

def merge_dataframes(df_primary: pl.DataFrame, df_secondary_mapped: pl.DataFrame, df_memofichas_mapped: pl.DataFrame | None = None) -> tuple[pl.DataFrame, dict[str, int]]:
    """
    Une las tres fuentes para basegeneralcostos con la misma regla:
    - BASE_GENERAL (primario): toda la base, fuente='BASE_GENERAL'.
    - POSTVENTA y MEMOFICHAS: solo se cargan las ordenes que NO estan ya en la base general;
      cada fila lleva en 'fuente' de donde provino ('POSTVENTA' o 'MEMOFICHAS').
    Prioridad si una orden aparece en varias fuentes: BASE_GENERAL > POSTVENTA > MEMOFICHAS.
    """
    log(f"[MERGE] Combinando DataFrames...")
    
    ordenes_ya_incluidas = set(df_primary.select("orden").to_series().cast(pl.Utf8).to_list())
    log(f"  Ordenes en primario (BASE_GENERAL): {len(ordenes_ya_incluidas)}")
    
    # POSTVENTA: solo ordenes que NO estan en la base general; fuente = POSTVENTA
    df_secondary_filtered = df_secondary_mapped.filter(
        ~pl.col("orden").cast(pl.Utf8).is_in(list(ordenes_ya_incluidas))
    )
    log(f"  Ordenes nuevas en POSTVENTA (se cargan con fuente=POSTVENTA): {df_secondary_filtered.height}")
    
    listas = [df_primary]
    if df_secondary_filtered.height > 0:
        listas.append(df_secondary_filtered)
        ordenes_ya_incluidas.update(df_secondary_filtered.select("orden").to_series().cast(pl.Utf8).to_list())
    
    # MEMOFICHAS: misma logica que POSTVENTA - solo ordenes que aun no estan en la base general; fuente = MEMOFICHAS
    if df_memofichas_mapped is not None and df_memofichas_mapped.height > 0:
        df_memofichas_filtered = df_memofichas_mapped.filter(
            ~pl.col("orden").cast(pl.Utf8).is_in(list(ordenes_ya_incluidas))
        )
        log(f"  Ordenes nuevas en MEMOFICHAS (se cargan con fuente=MEMOFICHAS): {df_memofichas_filtered.height}")
        if df_memofichas_filtered.height > 0:
            listas.append(df_memofichas_filtered)
    
    n_postventa = df_secondary_filtered.height
    n_memofichas = listas[2].height if len(listas) > 2 else 0
    counts = {"BASE_GENERAL": df_primary.height, "POSTVENTA": n_postventa, "MEMOFICHAS": n_memofichas}
    
    if len(listas) == 1:
        log("  No hay ordenes nuevas para agregar")
        return df_primary, counts
    
    # Misma lista de columnas en todos
    common_cols = list(df_primary.columns)
    for d in listas[1:]:
        common_cols = [c for c in common_cols if c in d.columns]
    
    df_merged = pl.concat([d.select(common_cols) for d in listas])
    
    log(f"  Total filas despues del merge: {df_merged.height}")
    log(f"    - Del primario (BASE_GENERAL): {df_primary.height}")
    log(f"    - Del secundario (POSTVENTA): {n_postventa}")
    if len(listas) > 2:
        log(f"    - De MEMOFICHAS: {n_memofichas}")
    
    return df_merged, counts


# ==============================================================================
# SUBIDA A POSTGRESQL
# ==============================================================================

def upload_to_postgres(df: pl.DataFrame):
    """Sube el DataFrame a PostgreSQL"""
    log(f"[UPLOAD] Subiendo datos a PostgreSQL...")
    log(f"  Tabla destino: {TABLE_NAME_PG}")
    log(f"  Filas a insertar: {df.height}")
    log(f"  Columnas: {df.width}")
    
    t0 = time.time()
    
    try:
        # Asegurar codificacion UTF-8 correcta antes de escribir
        df = ensure_utf8_encoding(df)
        log(f"  Codificacion UTF-8 normalizada")
        
        # Reemplazar valores NULL: texto con string vacio, numericos con 0
        df = fill_null_values(df)
        
        # Primero vaciar la tabla (TRUNCATE) para no romper las vistas dependientes
        engine = create_engine(DB_URI)
        with engine.connect() as conn:
            conn.execute(text(f"TRUNCATE TABLE {TABLE_NAME_PG}"))
            conn.commit()
        log(f"  Tabla vaciada (TRUNCATE) - Vistas preservadas")
        
        # Luego insertar los datos (append)
        df.write_database(
            table_name=TABLE_NAME_PG, 
            connection=DB_URI, 
            if_table_exists="append",
            engine="adbc"
        )
        log(f"  Usando motor ADBC (Ultra Rapido)")
        log(f"  Carga completada en {time.time() - t0:.2f}s")
        
    except Exception as adbc_error:
        log(f"  Error ADBC: {adbc_error}", "error")
        log("  Verifica que PostgreSQL este activo en el puerto configurado.", "error")
        raise

