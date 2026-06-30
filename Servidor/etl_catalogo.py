import logging
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from logging.handlers import RotatingFileHandler

import fastexcel
import polars as pl
from sqlalchemy import create_engine, text

from etl_common import (
    CATALOGO_COLUMN_MAPPING,
    CATALOGO_SCHEMA,
    DB_URI,
    FILE_CATALOGO,
    HEADER_ROW_CATALOGO,
    SHEET_CATALOGO,
    TABLE_CATALOGO,
    clean_and_deduplicate_headers,
    ensure_utf8_encoding,
    fill_null_values,
    find_header_row,
    log,
    normalize_column_name,
    set_logger,
)
from etl_ot import (
    map_memofichas_to_schema,
    map_secondary_to_schema,
    merge_dataframes,
    read_memofichas_excel,
    read_primary_excel,
    read_secondary_excel,
    upload_to_postgres,
)

# ------------------------------------------------------------------------------
# PROVEEDORES (hoja proveedores -> tabla proveedorsolicitud)
# ------------------------------------------------------------------------------
SHEET_PROVEEDORES = "PROVEEDORES"
TABLE_PROVEEDORES = "proveedorsolicitud"
PROVEEDORES_SCHEMA = ["nit", "nombre"]
# La hoja PROVEEDORES no trae encabezados en la lectura fastexcel:
# col 0 = NIT, col 1 = nombre (PROVEEDORES en Excel); demas columnas se ignoran.
PROVEEDORES_NIT_COL = 0
PROVEEDORES_NOMBRE_COL = 1

# ==============================================================================
# LECTURA DE CATALOGO DE PRODUCTOS (para pipeline paralelo)
# ==============================================================================

def read_catalogo_excel() -> pl.DataFrame:
    """Lee el catalogo de productos desde Excel y lo procesa"""
    log(f"[CATALOGO] Leyendo: {FILE_CATALOGO}")
    log(f"           Hoja: {SHEET_CATALOGO}")
    
    t0 = time.time()
    
    # Leer Excel optimizado usando fastexcel directamente
    excel = fastexcel.read_excel(FILE_CATALOGO)
    df = excel.load_sheet_by_name(SHEET_CATALOGO).to_polars()
    # Buscar cabeceras dinamicamente
    keywords = ["REFERENCIA", "LIN", "GRU", "DESC"]
    found_offset = find_header_row(df, keywords)
    
    if found_offset != -1:
        raw_headers = df.row(found_offset)
        final_headers = clean_and_deduplicate_headers(raw_headers)
        df = df.slice(found_offset + 1)
        df.columns = final_headers
        log(f"  Cabeceras encontradas en indice {found_offset}")
    else:
        # Fallback a la fila configurada si no se encuentran keywords
        header_idx = HEADER_ROW_CATALOGO - 1
        if header_idx < df.height:
            raw_headers = df.row(header_idx)
            final_headers = clean_and_deduplicate_headers(raw_headers)
            df = df.slice(header_idx + 1)
            df.columns = final_headers
            log(f"  Cabeceras tomadas de fila configurada {HEADER_ROW_CATALOGO}")
    
    # Normalizar nombres de columnas
    df.columns = [normalize_column_name(c) for c in df.columns]
    log(f"  Columnas normalizadas: {df.columns}")
    
    # Aplicar mapeo de columnas
    if CATALOGO_COLUMN_MAPPING:
        rename_dict = {k: v for k, v in CATALOGO_COLUMN_MAPPING.items() if k in df.columns}
        if rename_dict:
            df = df.rename(rename_dict)
            log(f"  Columnas mapeadas: {len(rename_dict)}")
    
    # Deduplicar por referencia de forma agresiva
    if 'referencia' in df.columns:
        before = df.height
        # Normalizar para deduccion: strip y upper
        df = df.with_columns(pl.col('referencia').cast(pl.Utf8).str.strip_chars().str.to_uppercase().alias('referencia'))
        df = df.unique(subset=['referencia'], keep='first')
        # Eliminar filas donde referencia sea vacia
        df = df.filter(pl.col('referencia').str.len_chars() > 0)
        if df.height < before:
            log(f"  [DEDUPLICACION] Eliminados {before - df.height} registros duplicados o invalidos de Catalogo")

    # Agregar fecha y hora actuales (son NOT NULL en la BD)
    # Como fecha y hora son TEXT en PostgreSQL, usamos strings directamente
    now = datetime.now()
    df = df.with_columns([
        pl.lit(now.strftime('%Y-%m-%d')).alias('fecha'),
        pl.lit(now.strftime('%H:%M:%S')).alias('hora')
    ])
    
    # Agregar columnas faltantes como NULL (con tipo String para compatibilidad ADBC)
    for col in CATALOGO_SCHEMA:
        if col not in df.columns:
            df = df.with_columns(pl.lit(None).cast(pl.Utf8).alias(col))
    
    # Reordenar segun esquema
    df = df.select(CATALOGO_SCHEMA)
    
    # Castear columnas según tipo en BD
    CATALOGO_FLOAT_COLS = {'capacidad'}
    text_cols = [c for c in CATALOGO_SCHEMA if c not in ['fecha', 'hora'] and c not in CATALOGO_FLOAT_COLS]
    for col in text_cols:
        df = df.with_columns(pl.col(col).cast(pl.Utf8))
    for col in CATALOGO_FLOAT_COLS:
        if col in df.columns:
            df = df.with_columns(pl.col(col).cast(pl.Float64, strict=False))
    
    log(f"  Lectura completada en {time.time() - t0:.2f}s")
    log(f"  Filas: {df.height}")
    
    return df


# ==============================================================================
# LECTURA DE PROVEEDORES (hoja proveedores en CATALOGO FINAL)
# ==============================================================================

def read_proveedores_excel() -> pl.DataFrame:
    """Lee la hoja PROVEEDORES y retorna solo NIT y nombre."""
    log(f"[PROVEEDORES] Leyendo: {FILE_CATALOGO}")
    log(f"              Hoja: {SHEET_PROVEEDORES}")

    t0 = time.time()

    excel = fastexcel.read_excel(FILE_CATALOGO)
    df = excel.load_sheet_by_name(SHEET_PROVEEDORES).to_polars()

    if df.width < 2:
        raise ValueError(
            f"La hoja {SHEET_PROVEEDORES} debe tener al menos 2 columnas (NIT y nombre)"
        )

    nit_col = df.columns[PROVEEDORES_NIT_COL]
    nombre_col = df.columns[PROVEEDORES_NOMBRE_COL]
    df = df.select([nit_col, nombre_col])
    df = df.rename({nit_col: "nit", nombre_col: "nombre"})
    log(f"  Columnas usadas: posicion {PROVEEDORES_NIT_COL}=NIT, {PROVEEDORES_NOMBRE_COL}=nombre")

    before = df.height
    df = df.with_columns([
        pl.col("nit")
        .cast(pl.Float64, strict=False)
        .round(0)
        .cast(pl.Int64, strict=False)
        .cast(pl.Utf8, strict=False)
        .fill_null("")
        .str.strip_chars()
        .alias("nit"),
        pl.col("nombre")
        .cast(pl.Utf8, strict=False)
        .fill_null("")
        .str.strip_chars()
        .alias("nombre"),
    ])
    df = df.filter(pl.col("nit").str.len_chars() > 0)
    df = df.filter(pl.col("nit").str.to_uppercase() != "NIT")
    df = df.unique(subset=["nit"], keep="first")
    df = df.unique(subset=["nombre"], keep="first")
    if df.height < before:
        log(f"  [DEDUPLICACION] Eliminados {before - df.height} registros duplicados o invalidos de Proveedores")

    df = df.select(PROVEEDORES_SCHEMA)
    for col in PROVEEDORES_SCHEMA:
        df = df.with_columns(pl.col(col).cast(pl.Utf8))

    log(f"  Lectura completada en {time.time() - t0:.2f}s")
    log(f"  Filas: {df.height}")

    return df


# ==============================================================================
# SUBIDA DE CATALOGO A POSTGRESQL
# ==============================================================================

def upload_catalogo_to_postgres(df: pl.DataFrame):
    """Sube el DataFrame del catalogo a PostgreSQL usando ADBC (igual que OT)"""
    log(f"[UPLOAD CATALOGO] Subiendo a tabla {TABLE_CATALOGO}...")
    log(f"  Filas a insertar: {df.height}")
    
    t0 = time.time()
    
    try:
        # Asegurar codificacion UTF-8 correcta antes de escribir
        df = ensure_utf8_encoding(df)
        log(f"  Codificacion UTF-8 normalizada")
        
        # Reemplazar valores NULL: texto con string vacio, numericos con 0
        df = fill_null_values(df)
        
        # TRUNCATE primero
        engine = create_engine(DB_URI)
        with engine.connect() as conn:
            conn.execute(text(f"TRUNCATE TABLE {TABLE_CATALOGO}"))
            conn.commit()
        log(f"  Tabla vaciada (TRUNCATE)")
        
        # Usar ADBC (igual que para OT)
        df.write_database(
            table_name=TABLE_CATALOGO,
            connection=DB_URI,
            if_table_exists="append",
            engine="adbc"
        )
        
        log(f"  Usando motor ADBC (Ultra Rapido)")
        log(f"  Carga completada en {time.time() - t0:.2f}s")
        return df.height
        
    except Exception as e:
        log(f"  Error ADBC: {e}", "warning")
        log(f"  Intentando fallback con INSERT por lotes...")
        # El fallback debe re-truncar porque ADBC pudo haber dejado basura
        with engine.connect() as conn:
            conn.execute(text(f"TRUNCATE TABLE {TABLE_CATALOGO}"))
            conn.commit()
        return upload_catalogo_fallback(df)


def upload_catalogo_fallback(df: pl.DataFrame):
    """Fallback usando INSERT por lotes si ADBC falla"""
    t0 = time.time()
    
    # Aplicar las mismas transformaciones que en la funcion principal
    df = ensure_utf8_encoding(df)
    df = fill_null_values(df)
    
    df_upload = df.with_columns([
        pl.col('fecha').cast(pl.Utf8),
        pl.col('hora').cast(pl.Utf8)
    ])
    
    records = df_upload.to_dicts()
    columns = ', '.join(CATALOGO_SCHEMA)
    placeholders = ', '.join([f':{col}' for col in CATALOGO_SCHEMA])
    insert_sql = text(f"INSERT INTO {TABLE_CATALOGO} ({columns}) VALUES ({placeholders})")
    
    engine = create_engine(DB_URI)
    batch_size = 1000
    with engine.connect() as conn:
        for i in range(0, len(records), batch_size):
            batch = records[i:i+batch_size]
            conn.execute(insert_sql, batch)
        conn.commit()
    
    log(f"  Carga completada en {time.time() - t0:.2f}s (INSERT fallback)")
    return df.height


# ==============================================================================
# SUBIDA DE PROVEEDORES A POSTGRESQL
# ==============================================================================

def upload_proveedores_to_postgres(df: pl.DataFrame):
    """Sube proveedores a PostgreSQL usando ADBC (mismo patron que catalogo)."""
    log(f"[UPLOAD PROVEEDORES] Subiendo a tabla {TABLE_PROVEEDORES}...")
    log(f"  Filas a insertar: {df.height}")

    t0 = time.time()
    engine = create_engine(DB_URI)

    try:
        df = ensure_utf8_encoding(df)
        log(f"  Codificacion UTF-8 normalizada")

        df = fill_null_values(df)

        before = df.height
        df = df.unique(subset=["nit"], keep="first")
        df = df.unique(subset=["nombre"], keep="first")
        if df.height < before:
            log(f"  [DEDUPLICACION POST-LIMPIEZA] Eliminados {before - df.height} proveedores duplicados")

        with engine.connect() as conn:
            conn.execute(text(f"TRUNCATE TABLE {TABLE_PROVEEDORES}"))
            conn.commit()
        log(f"  Tabla vaciada (TRUNCATE)")

        df.write_database(
            table_name=TABLE_PROVEEDORES,
            connection=DB_URI,
            if_table_exists="append",
            engine="adbc",
        )

        log(f"  Usando motor ADBC (Ultra Rapido)")
        log(f"  Carga completada en {time.time() - t0:.2f}s")
        return df.height

    except Exception as e:
        log(f"  Error ADBC: {e}", "warning")
        log(f"  Intentando fallback con INSERT por lotes...")
        with engine.connect() as conn:
            conn.execute(text(f"TRUNCATE TABLE {TABLE_PROVEEDORES}"))
            conn.commit()
        return upload_proveedores_fallback(df)


def upload_proveedores_fallback(df: pl.DataFrame):
    """Fallback usando INSERT por lotes si ADBC falla."""
    t0 = time.time()

    df = ensure_utf8_encoding(df)
    df = fill_null_values(df)

    before = df.height
    df = df.unique(subset=["nit"], keep="first")
    df = df.unique(subset=["nombre"], keep="first")
    if df.height < before:
        log(f"  [DEDUPLICACION POST-LIMPIEZA] Eliminados {before - df.height} proveedores duplicados")

    records = df.to_dicts()
    columns = "nit, nombre"
    placeholders = ":nit, :nombre"
    insert_sql = text(f"INSERT INTO {TABLE_PROVEEDORES} ({columns}) VALUES ({placeholders})")

    engine = create_engine(DB_URI)
    batch_size = 1000
    with engine.connect() as conn:
        for i in range(0, len(records), batch_size):
            batch = records[i:i + batch_size]
            conn.execute(insert_sql, batch)
        conn.commit()

    log(f"  Carga completada en {time.time() - t0:.2f}s (INSERT fallback)")
    return df.height


def upload_catalogo():
    """Carga catalogo de productos y proveedores (endpoint /sync/catalogo)."""
    log("=" * 70)
    log("  CARGA CATALOGO Y PROVEEDORES (STANDALONE)")
    log("=" * 70)

    start_time = time.time()

    try:
        df_cat = read_catalogo_excel()
        registros_cat = upload_catalogo_to_postgres(df_cat)

        df_prov = read_proveedores_excel()
        registros_prov = upload_proveedores_to_postgres(df_prov)

        elapsed = time.time() - start_time
        log("=" * 70)
        log(f"  CATALOGO COMPLETADO - {registros_cat} registros")
        log(f"  PROVEEDORES COMPLETADO - {registros_prov} registros")
        log(f"  Tiempo total: {elapsed:.2f}s")
        log("=" * 70)

        return {"catalogo": registros_cat, "proveedores": registros_prov}

    except Exception as e:
        log(f"ERROR en carga de catalogo/proveedores: {e}", "error")
        raise

# ==============================================================================
# LECTURA PARALELA (4 ARCHIVOS SIMULTANEOS)
# ==============================================================================

def read_files_parallel(include_catalogo: bool = True):
    """Lee y procesa los archivos en paralelo.
    Si include_catalogo=True: 4 archivos (primario, secundario, memofichas, catalogo).
    Si include_catalogo=False: solo 3 archivos para basegeneralcostos (sin catalogo).
    """
    if include_catalogo:
        log(f"[PARALELO] Iniciando pipeline de procesamiento (4 hilos)...")
        workers = 4
    else:
        log(f"[PARALELO] Iniciando pipeline de procesamiento (3 hilos, sin catalogo)...")
        workers = 3
    t0 = time.time()

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_primary = executor.submit(read_primary_excel)

        def secondary_pipeline():
            return map_secondary_to_schema(read_secondary_excel())
        future_secondary_mapped = executor.submit(secondary_pipeline)

        def memofichas_pipeline():
            df = read_memofichas_excel()
            return map_memofichas_to_schema(df) if df is not None else None
        future_memofichas_mapped = executor.submit(memofichas_pipeline)

        df_primary = future_primary.result()
        df_secondary_mapped = future_secondary_mapped.result()
        df_memofichas_mapped = future_memofichas_mapped.result()

        if include_catalogo:
            future_catalogo = executor.submit(read_catalogo_excel)
            df_catalogo = future_catalogo.result()
        else:
            df_catalogo = None

    elapsed = time.time() - t0
    log(f"[PARALELO] Pipeline completado en {elapsed:.2f}s ({workers} archivos)")

    return df_primary, df_secondary_mapped, df_memofichas_mapped, df_catalogo


# ==============================================================================
# FUNCION PRINCIPAL
# ==============================================================================

def upload_buffer_with_merge(include_catalogo: bool = False):
    """Funcion principal: Pipeline paralelo, merge (primario+secundario+MEMOFICHAS) y upload.
    Por defecto include_catalogo=False: solo carga basegeneralcostos (OT).
    Si include_catalogo=True: ademas lee y sube el catalogo en paralelo (comportamiento legacy).
    Para cargar solo el catalogo, usar el endpoint /sync/catalogo que llama a upload_catalogo().
    """
    log("=" * 70)
    if include_catalogo:
        log("  CARGA COMPLETA (PRIMARIO + POSTVENTA + MEMOFICHAS + CATALOGO)")
    else:
        log("  CARGA BASE GENERAL COSTOS (PRIMARIO + POSTVENTA + MEMOFICHAS)")
    log("=" * 70)

    start_time = time.time()

    try:
        # 1. Ejecutar el pipeline paralelo (3 o 4 archivos segun include_catalogo)
        df_primary, df_secondary_mapped, df_memofichas_mapped, df_catalogo = read_files_parallel(include_catalogo=include_catalogo)

        # 2. Merge y upload de basegeneralcostos; opcionalmente catalogo en paralelo
        if include_catalogo and df_catalogo is not None:
            with ThreadPoolExecutor(max_workers=2) as executor:
                def ot_upload_task():
                    df_merged, counts = merge_dataframes(df_primary, df_secondary_mapped, df_memofichas_mapped)
                    upload_to_postgres(df_merged)
                    return df_merged.height, counts

                def cat_upload_task():
                    cat_count = upload_catalogo_to_postgres(df_catalogo)
                    df_prov = read_proveedores_excel()
                    prov_count = upload_proveedores_to_postgres(df_prov)
                    return cat_count, prov_count

                future_ot = executor.submit(ot_upload_task)
                future_cat = executor.submit(cat_upload_task)

                total_ot, counts_ot = future_ot.result()
                total_cat, total_prov = future_cat.result()
        else:
            df_merged, counts_ot = merge_dataframes(df_primary, df_secondary_mapped, df_memofichas_mapped)
            upload_to_postgres(df_merged)
            total_ot = df_merged.height
            total_cat = None
            total_prov = None

        elapsed = time.time() - start_time
        log("=" * 70)
        log("  COMPLETADO EXITOSAMENTE")
        log(f"  - basegeneralcostos: {total_ot} registros totales")
        log(f"    - BASE_GENERAL: {counts_ot['BASE_GENERAL']} | POSTVENTA: {counts_ot['POSTVENTA']} | MEMOFICHAS: {counts_ot['MEMOFICHAS']}")
        if total_cat is not None:
            log(f"  - Catalogo: {total_cat} registros")
        if total_prov is not None:
            log(f"  - Proveedores: {total_prov} registros")
        log(f"  Tiempo total: {elapsed:.2f} segundos")
        log("=" * 70)

    except Exception as e:
        log(f"ERROR FATAL: {e}", "error")
        raise


if __name__ == "__main__":
    basedir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(basedir, "sync.log")

    _logger = logging.getLogger("SyncOT")
    _logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "[%(asctime)s] %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler = RotatingFileHandler(
        log_file, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    _logger.addHandler(file_handler)
    set_logger(_logger)

    try:
        upload_buffer_with_merge(include_catalogo=True)
        sys.exit(0)
    except Exception as e:
        log(f"ERROR FATAL: {e}", "error")
        sys.exit(1)

