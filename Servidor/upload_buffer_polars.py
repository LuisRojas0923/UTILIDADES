import os
import polars as pl
import fastexcel
import time
import re
import logging
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from sqlalchemy import create_engine, text

# ==============================================================================
# LOGGER (compartido con api_server.py o standalone)
# ==============================================================================
_logger = None

def set_logger(logger):
    """Recibe el logger desde api_server.py"""
    global _logger
    _logger = logger

def log(message, level="info"):
    """Funcion de logging que funciona con o sin logger externo"""
    global _logger
    if _logger:
        if level == "error":
            _logger.error(message)
        elif level == "warning":
            _logger.warning(message)
        else:
            _logger.info(message)
    # Sin fallback a consola - solo loguea si hay logger configurado


# ==============================================================================
# CONFIGURACION DEL SCRIPT
# ==============================================================================

# Conexion a PostgreSQL (override: DB_NAME o DB_URI completas para pruebas)
_DB_NAME = os.environ.get("DB_NAME", "solid").strip()
DB_URI = os.environ.get(
    "DB_URI",
    f"postgresql://postgres:AdminSolid2025@192.168.0.21:5432/{_DB_NAME}",
).strip()

# Tabla destino en PostgreSQL
TABLE_NAME_PG = "basegeneralcostos"

# ------------------------------------------------------------------------------
# RUTAS DE ARCHIVOS EXCEL
# Si EXCEL_BASE_PATH está definida (ej. en Docker/Linux con SMB montado), se usan
# rutas bajo esa carpeta. Si no, se usan rutas UNC de Windows (\\192.168.0.3\...).
# ------------------------------------------------------------------------------
_EXCEL_BASE = os.environ.get("EXCEL_BASE_PATH", "").strip()
if _EXCEL_BASE:
    FILE_PRIMARY = os.path.join(_EXCEL_BASE, "Procesos Comunes SGI", "Costos", "INFORME DE ORDENES", "BASE DE DATOS GENERAL.xlsm")
    FILE_SECONDARY = os.path.join(_EXCEL_BASE, "Postventa", "MANTENIMIENTO Y SERVICIO POSTVENTA", "- GESTION ORDENES DE SERVICIO", "CENTRO LOGÍSTICO", "MTZ-SPT-02 Informe Gestion Postventa V1.xlsm")
    FILE_CATALOGO = os.path.join(_EXCEL_BASE, "Procesos Comunes SGI", "Mejora", "Catalogo de Articulos", "CATALOGO FINAL.xlsx")
    FILE_MEMOFICHAS = os.path.join(_EXCEL_BASE, "Control Presupuestal", "MEMOFICHA", "CONSULTAS", "CONSULTA MEMOFICHAS v2.xlsx")
    FILE_PROVEEDORES = os.path.join(_EXCEL_BASE, "Control Presupuestal", "CATALOGO DE PRODUCTOS", "CATALOGO DE PRODUCTOS.xlsm")
else:
    FILE_PRIMARY = r"\\192.168.0.3\Procesos Comunes SGI\Costos\INFORME DE ORDENES\BASE DE DATOS GENERAL.xlsm"
    FILE_SECONDARY = r"\\192.168.0.3\Postventa\MANTENIMIENTO Y SERVICIO POSTVENTA\- GESTION ORDENES DE SERVICIO\CENTRO LOGÍSTICO\MTZ-SPT-02 Informe Gestion Postventa V1.xlsm"
    FILE_CATALOGO = r"\\192.168.0.3\Procesos Comunes SGI\Mejora\Catalogo de Articulos\CATALOGO FINAL.xlsx"
    FILE_MEMOFICHAS = r"\\192.168.0.3\Control Presupuestal\MEMOFICHA\CONSULTAS\CONSULTA MEMOFICHAS v2.xlsx"
    FILE_PROVEEDORES = r"\\192.168.0.3\Control Presupuestal\CATALOGO DE PRODUCTOS\CATALOGO DE PRODUCTOS.xlsm"

# ------------------------------------------------------------------------------
# ARCHIVO PRIMARIO (Base General de Costos)
# ------------------------------------------------------------------------------
SHEET_PRIMARY = "MOVIMIENTO_DE ORDENES__2"

# ------------------------------------------------------------------------------
# ARCHIVO SECUNDARIO (Informe Gestion Postventa)
# ------------------------------------------------------------------------------
SHEET_SECONDARY = "Info_GesPostV"
HEADER_ROW_SECONDARY = 3  # Fila donde estan los encabezados (1-indexed, fila 3 en Excel)

# ------------------------------------------------------------------------------
# ARCHIVO MEMOFICHAS (Control Presupuestal - se une a basegeneralcostos)
# ------------------------------------------------------------------------------
SHEET_MEMOFICHAS = "MEMOFICHAS"
# Cabeceras: OP, No., ORDEN, SUBINDICE, TIPO DE B, DESCRIPCION PRODUCTO TERMINADO, UND., CANTIDAD, CODIGO PT, CODIGO PP, CC, SCC, UEN PN, ESP, CLIENTE, INGENIERO, B DESCRIPCION, CONCEPTO, SUBINDICE_1, SUBINDICE_2
HEADER_ROW_MEMOFICHAS = 0  # Se detecta con find_header_row (keywords: ORDEN, OP, CC)

# ------------------------------------------------------------------------------
# ARCHIVO CATALOGO DE PRODUCTOS
# ------------------------------------------------------------------------------
SHEET_CATALOGO = "CATALOGO SIIGO"
HEADER_ROW_CATALOGO = 4  # Fila donde estan los encabezados (1-indexed): REFERENCIA, LIN, GRU, etc.
TABLE_CATALOGO = "catalogoproducto"

# Esquema de la tabla catalogoproducto (columnas en el orden de la BD)
CATALOGO_SCHEMA = [
    'referencia', 'fecha', 'hora', 'codigolinea', 'codigogrupo', 'elemento',
    'descripcion', 'unidadmedida', 'linea', 'grupo', 'tipo', 'clasificacion',
    'rotacion', 'periodo', 'proveedorfrecuente', 'clasificacioncompras', 'formato',
    'capacidad', 'ubicacionalmacen'
]

# Mapeo de columnas Excel -> BD
# Clave: nombre normalizado en Excel, Valor: nombre en la BD
CATALOGO_COLUMN_MAPPING = {
    'lin': 'codigolinea',
    'gru': 'codigogrupo',
    'unidad': 'unidadmedida',
    'nal_imp': 'tipo',
    'nivel_de_rotacion': 'rotacion',
    'periodo_actual_clasificaci_n': 'periodo',  # En el Excel viene con guion bajo
    'proveedor_frecuente': 'proveedorfrecuente',
    'clasificacion_compras': 'clasificacioncompras',
    'tipo_de_formato': 'formato',
    'ubicacion_almacen': 'ubicacionalmacen',
}

# ------------------------------------------------------------------------------
# ARCHIVO PROVEEDORES (hoja PROVEEDOR PRINC del catalogo de productos)
# Campos alineados con tabla ERP proveedor: nit, nombre
# ------------------------------------------------------------------------------
SHEET_PROVEEDORES = "PROVEEDOR PRINC"
TABLE_PROVEEDORES = "proveedorprincip"
PROVEEDOR_SCHEMA = ["nit", "nombre"]
PROVEEDOR_COLUMN_MAPPING = {
    "proveedor": "nombre",
}

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
    'ciudad': 'ubicacion',                     # Posicion 5
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

# ------------------------------------------------------------------------------
# MAPEO DE COLUMNAS: MEMOFICHAS -> Esquema basegeneralcostos
# ------------------------------------------------------------------------------
# Nombres del Excel (normalizados a snake_case) -> columnas FULL_SCHEMA
COLUMN_MAPPING_MEMOFICHAS = {
    'op': 'op',
    'orden': 'orden',
    'subindice': 'sub_indice',
    'tipo_de_b': 'b',
    'descripcion_producto_terminado': 'descripcion_producto_terminado',
    'codigo_pt': 'codigo_pt',
    'codigo_pp': 'codigo_pp',
    'cc': 'cc',
    'scc': 'scc',
    'uen_pn': 'uen',
    'esp': 'especialidad',
    'cliente': 'cliente',
    'ingeniero': 'ingeniero',
    'b_descripcion': 'descripcion',
    'subindice_1': 'categoria_sub_indice',
}

# Columnas del esquema que no existen en MEMOFICHAS (seran NULL)
NULL_COLUMNS_FOR_MEMOFICHAS = [
    'orden_vieja_nueva', 'clasificacion_de_la_orden', 'cod_uen',
    'estado', 'apertura_siigo', 'cierre_siigo', 'estado_base_contrato',
    'apertura_base_contratos', 'cierre_base_contratos', 'ot_planta', 'nit', 'uen_fact',
    'vr_contratado', 'ubicacion', 'cajas_menores'
]

# Esquema completo (30 columnas originales + 1 fuente)
FULL_SCHEMA = [
    'op', 'orden', 'orden_vieja_nueva', 'clasificacion_de_la_orden', 'cc', 'scc', 'b', 
    'especialidad', 'cod_uen', 'uen', 'sub_indice', 'estado', 'apertura_siigo', 
    'cierre_siigo', 'descripcion_producto_terminado', 'codigo_pp', 'codigo_pt', 
    'categoria_sub_indice', 'cliente', 'ubicacion', 'estado_base_contrato', 
    'apertura_base_contratos', 'cierre_base_contratos', 'ingeniero', 'ot_planta', 
    'nit', 'uen_fact', 'vr_contratado', 'descripcion', 'cajas_menores', 'fuente'
]


# ==============================================================================
# FUNCIONES DE UTILIDAD
# ==============================================================================

def normalize_column_name(col: str) -> str:
    """Normaliza un nombre de columna a snake_case sin acentos"""
    # 1. Mapeo explícito de tildes para mayor robustez
    tab = str.maketrans(
        "áéíóúüñÁÉÍÓÚÜÑ",
        "aeiouunAEIOUUN"
    )
    clean_col = col.translate(tab)
    
    # 2. Por si acaso, normalización NFD para otros caracteres
    clean_col = "".join(
        c for c in unicodedata.normalize('NFD', clean_col)
        if unicodedata.category(c) != 'Mn'
    )
    
    # 3. Convertir a minusculas y snake_case
    clean_col = clean_col.strip().lower()
    clean_col = re.sub(r'[^a-z0-9]+', '_', clean_col)
    clean_col = clean_col.strip('_')
    return clean_col


def remove_accents_vectorized(col_expr: pl.Expr) -> pl.Expr:
    """Versión vectorizada de limpieza de caracteres para Polars.
    Mucho más rápida que map_elements ya que usa expresiones nativas.
    """
    # 1. Normalización básica de tildes y eñes
    expr = (
        col_expr.cast(pl.String)
        .str.replace_all(r"[áàäâÁÀÄÂ]", "A")
        .str.replace_all(r"[éèëêÉÈËÊ]", "E")
        .str.replace_all(r"[íìïîÍÌÏÎ]", "I")
        .str.replace_all(r"[óòöôÓÒÖÔ]", "O")
        .str.replace_all(r"[úùüûÚÙÜÛ]", "U")
        .str.replace_all(r"[ñÑ]", "N")
        # 2. Corregir errores comunes de doble codificación (Ã‰ -> E, etc)
        .str.replace_all(r"Ã[‰]", "E")
        .str.replace_all(r"Ã[ ]", "A")
        .str.replace_all(r"Ã[í]", "I")
        .str.replace_all(r"Ã³", "O")
        .str.replace_all(r"Ãº", "U")
        .str.replace_all(r"Ã±", "N")
        # 3. Eliminar cualquier carácter NO ASCII o de control
        # Rango [^ -~] incluye todo lo que no sea [32-126] en ASCII
        .str.replace_all(r"[^ -~]", "")
        # 4. Limpieza de espacios
        .str.strip_chars()
        .str.replace_all(r"\s+", " ")
    )
    return expr


def ensure_utf8_encoding(df: pl.DataFrame) -> pl.DataFrame:
    """Convierte todas las columnas de texto a ASCII puro
    
    IMPORTANTE: El ERP no soporta UTF-8 y tiene problemas con caracteres de Windows-1252
    (como el byte 0x90), por lo que convertimos todo a ASCII puro eliminando:
    - Todos los caracteres no-ASCII
    - Caracteres de control (incluyendo 0x90 de Windows-1252)
    - Solo mantenemos caracteres ASCII imprimibles (32-126)
    """
    # Identificar columnas de texto (Utf8, String, Object)
    text_columns = []
    for col in df.columns:
        dtype = df[col].dtype
        if dtype == pl.Utf8 or dtype == pl.String or dtype == pl.Object:
            text_columns.append(col)
    
    if not text_columns:
        return df
    
    log(f"  Convirtiendo {len(text_columns)} columnas de texto a ASCII puro (eliminando caracteres Windows-1252 problemáticos)")
    
    # Convertir todas las columnas de texto a ASCII puro
    # Aplicar todas las conversiones de forma vectorizada
    expressions = [
        remove_accents_vectorized(pl.col(col)).alias(col)
        for col in text_columns
    ]
    
    # Aplicar todas las conversiones
    df_encoded = df.with_columns(expressions)
    
    return df_encoded


def fill_null_values(df: pl.DataFrame) -> pl.DataFrame:
    """Reemplaza valores NULL: texto con string vacio, numericos con 0
    
    Esta funcion procesa TODAS las columnas y reemplaza cualquier NULL encontrado.
    """
    expressions = []
    text_count = 0
    numeric_count = 0
    other_count = 0
    
    for col in df.columns:
        dtype = df[col].dtype
        dtype_str = str(dtype)
        
        # Manejar tipos Null/Unknown convirtiendolos primero
        if dtype == pl.Null or dtype_str == 'Null' or 'Null' in dtype_str:
            # Si es tipo Null, convertirlo a texto por defecto
            expressions.append(pl.col(col).cast(pl.Utf8, strict=False).fill_null("").alias(col))
            text_count += 1
            continue
        
        # Identificar si es texto o numerico
        is_text = (dtype == pl.Utf8 or dtype == pl.String or dtype == pl.Object or 
                   'Utf8' in dtype_str or 'String' in dtype_str)
        is_numeric = (dtype in [
            pl.Int8, pl.Int16, pl.Int32, pl.Int64,
            pl.UInt8, pl.UInt16, pl.UInt32, pl.UInt64,
            pl.Float32, pl.Float64, pl.Decimal
        ] or 'Int' in dtype_str or 'Float' in dtype_str or 'Decimal' in dtype_str)
        
        if is_text:
            # Reemplazar NULL con string vacio para columnas de texto
            # Asegurar que sea Utf8 primero, luego reemplazar NULLs
            expressions.append(
                pl.col(col)
                .cast(pl.Utf8, strict=False)
                .fill_null("")
                .alias(col)
            )
            text_count += 1
        elif is_numeric:
            # Reemplazar NULL con 0 para columnas numericas
            # Mantener el tipo original pero reemplazar NULLs
            expressions.append(
                pl.col(col)
                .fill_null(0)
                .alias(col)
            )
            numeric_count += 1
        else:
            # Para otros tipos (Date, Datetime, Boolean, etc.), convertir a texto y reemplazar
            # Esto asegura que no queden NULLs sin procesar
            expressions.append(
                pl.col(col)
                .cast(pl.Utf8, strict=False)
                .fill_null("")
                .alias(col)
            )
            other_count += 1
    
    if expressions:
        df_filled = df.with_columns(expressions)
        
        # Verificar que no queden NULLs - hacer una segunda pasada si es necesario
        null_counts = df_filled.null_count()
        total_nulls = null_counts.sum_horizontal().item()
        
        if total_nulls > 0:
            log(f"  ADVERTENCIA: Aun quedan {total_nulls} valores NULL despues del reemplazo", "warning")
            # Aplicar reemplazo agresivo a todas las columnas que aun tengan NULLs
            for col in df_filled.columns:
                null_count = null_counts[col].item()
                if null_count > 0:
                    dtype = df_filled[col].dtype
                    if dtype in [pl.Int8, pl.Int16, pl.Int32, pl.Int64, pl.UInt8, pl.UInt16, pl.UInt32, pl.UInt64, pl.Float32, pl.Float64]:
                        df_filled = df_filled.with_columns(pl.col(col).fill_null(0).alias(col))
                    else:
                        df_filled = df_filled.with_columns(pl.col(col).cast(pl.Utf8, strict=False).fill_null("").alias(col))
            log(f"  Reemplazo agresivo aplicado para eliminar NULLs restantes")
        
        if text_count > 0 or numeric_count > 0 or other_count > 0:
            log(f"  Valores NULL reemplazados: {text_count} columnas texto (-> ''), {numeric_count} columnas numericas (-> 0), {other_count} otros tipos (-> '')")
        
        return df_filled
    
    return df


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
        log(f"  Error ADBC: {adbc_error}", "warning")
        log(f"  Intentando fallback con INSERT por lotes...")
        upload_to_postgres_fallback(df)
        log(f"  Carga completada en {time.time() - t0:.2f}s (INSERT fallback)")


def upload_to_postgres_fallback(df: pl.DataFrame):
    """Fallback usando INSERT por lotes si ADBC falla"""
    df = ensure_utf8_encoding(df)
    df = fill_null_values(df)

    date_casts = [
        pl.col(col).cast(pl.Utf8).alias(col)
        for col in df.columns
        if df[col].dtype in (pl.Date, pl.Datetime)
    ]
    if date_casts:
        df = df.with_columns(date_casts)

    schema_cols = [c for c in FULL_SCHEMA if c in df.columns]
    records = df.select(schema_cols).to_dicts()
    columns = ", ".join(schema_cols)
    placeholders = ", ".join([f":{col}" for col in schema_cols])
    insert_sql = text(f"INSERT INTO {TABLE_NAME_PG} ({columns}) VALUES ({placeholders})")

    engine = create_engine(DB_URI)
    batch_size = 1000
    with engine.connect() as conn:
        for i in range(0, len(records), batch_size):
            batch = records[i : i + batch_size]
            conn.execute(insert_sql, batch)
        conn.commit()


# ==============================================================================
# LECTURA DE CATALOGO DE PRODUCTOS (para pipeline paralelo)
# ==============================================================================

def read_catalogo_excel() -> pl.DataFrame:
    """Lee el catalogo de productos desde Excel y lo procesa"""
    from datetime import datetime
    
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


def upload_catalogo():
    """Funcion standalone para cargar solo el catalogo (endpoint /catalogo)"""
    log("=" * 70)
    log("  CARGA CATALOGO DE PRODUCTOS (STANDALONE)")
    log("=" * 70)
    
    start_time = time.time()
    
    try:
        df = read_catalogo_excel()
        registros = upload_catalogo_to_postgres(df)
        
        elapsed = time.time() - start_time
        log("=" * 70)
        log(f"  CATALOGO COMPLETADO - {registros} registros en {elapsed:.2f}s")
        log("=" * 70)
        
        return registros
        
    except Exception as e:
        log(f"ERROR en carga de catalogo: {e}", "error")
        raise


# ==============================================================================
# LECTURA Y SUBIDA DE PROVEEDORES (hoja PROVEEDOR PRINC)
# ==============================================================================

def read_proveedores_excel() -> pl.DataFrame:
    """Lee proveedores unicos (nit, nombre) desde la hoja PROVEEDOR PRINC."""
    log(f"[PROVEEDORES] Leyendo: {FILE_PROVEEDORES}")
    log(f"              Hoja: {SHEET_PROVEEDORES}")

    t0 = time.time()
    excel = fastexcel.read_excel(FILE_PROVEEDORES)
    df = excel.load_sheet_by_name(SHEET_PROVEEDORES).to_polars()

    keywords = ["NIT", "PROVEEDOR", "REFERENCIA"]
    found_offset = find_header_row(df, keywords)

    if found_offset != -1:
        raw_headers = df.row(found_offset)
        final_headers = clean_and_deduplicate_headers(raw_headers)
        df = df.slice(found_offset + 1)
        df.columns = final_headers
        log(f"  Cabeceras encontradas en indice {found_offset}")
    else:
        raise ValueError("No se encontraron cabeceras NIT/PROVEEDOR en PROVEEDOR PRINC")

    df.columns = [normalize_column_name(c) for c in df.columns]
    log(f"  Columnas normalizadas: {df.columns}")

    rename_dict = {k: v for k, v in PROVEEDOR_COLUMN_MAPPING.items() if k in df.columns}
    if rename_dict:
        df = df.rename(rename_dict)
        log(f"  Columnas mapeadas: {rename_dict}")

    missing = [c for c in PROVEEDOR_SCHEMA if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas en PROVEEDOR PRINC: {missing}")

    before = df.height
    df = df.with_columns([
        pl.col("nit").cast(pl.Utf8).str.strip_chars().alias("nit"),
        pl.col("nombre").cast(pl.Utf8).str.strip_chars().alias("nombre"),
    ])
    df = df.filter((pl.col("nit").str.len_chars() > 0) & (pl.col("nombre").str.len_chars() > 0))
    df = df.unique(subset=["nit"], keep="first")
    df = df.select(PROVEEDOR_SCHEMA)

    if df.height < before:
        log(f"  [DEDUPLICACION] {before} filas -> {df.height} proveedores unicos por NIT")

    log(f"  Lectura completada en {time.time() - t0:.2f}s")
    log(f"  Proveedores unicos: {df.height}")
    return df


def ensure_proveedores_table():
    """Crea la tabla proveedorprincip si no existe (nit, nombre como en ERP proveedor)."""
    engine = create_engine(DB_URI)
    ddl = text(f"""
        CREATE TABLE IF NOT EXISTS {TABLE_PROVEEDORES} (
            nit text PRIMARY KEY,
            nombre text NOT NULL
        )
    """)
    with engine.connect() as conn:
        conn.execute(ddl)
        conn.commit()


def upload_proveedores_to_postgres(df: pl.DataFrame):
    """Sube proveedores unicos a PostgreSQL."""
    log(f"[UPLOAD PROVEEDORES] Subiendo a tabla {TABLE_PROVEEDORES}...")
    log(f"  Filas a insertar: {df.height}")

    ensure_proveedores_table()
    t0 = time.time()

    try:
        df = ensure_utf8_encoding(df)
        df = fill_null_values(df)

        engine = create_engine(DB_URI)
        with engine.connect() as conn:
            conn.execute(text(f"TRUNCATE TABLE {TABLE_PROVEEDORES}"))
            conn.commit()
        log("  Tabla vaciada (TRUNCATE)")

        df.write_database(
            table_name=TABLE_PROVEEDORES,
            connection=DB_URI,
            if_table_exists="append",
            engine="adbc",
        )
        log(f"  Usando motor ADBC")
        log(f"  Carga completada en {time.time() - t0:.2f}s")
        return df.height

    except Exception as e:
        log(f"  Error ADBC: {e}", "warning")
        log("  Intentando fallback con INSERT por lotes...")
        return upload_proveedores_fallback(df)


def upload_proveedores_fallback(df: pl.DataFrame):
    """Fallback usando INSERT por lotes si ADBC falla."""
    t0 = time.time()
    df = ensure_utf8_encoding(df)
    df = fill_null_values(df)

    records = df.select(PROVEEDOR_SCHEMA).to_dicts()
    columns = ", ".join(PROVEEDOR_SCHEMA)
    placeholders = ", ".join([f":{col}" for col in PROVEEDOR_SCHEMA])
    insert_sql = text(f"INSERT INTO {TABLE_PROVEEDORES} ({columns}) VALUES ({placeholders})")

    engine = create_engine(DB_URI)
    batch_size = 1000
    with engine.connect() as conn:
        for i in range(0, len(records), batch_size):
            batch = records[i : i + batch_size]
            conn.execute(insert_sql, batch)
        conn.commit()

    log(f"  Carga completada en {time.time() - t0:.2f}s (INSERT fallback)")
    return df.height


def upload_proveedores():
    """Funcion standalone para cargar proveedores (endpoint /sync/proveedores)."""
    log("=" * 70)
    log("  CARGA PROVEEDORES (PROVEEDOR PRINC)")
    log("=" * 70)

    start_time = time.time()
    try:
        df = read_proveedores_excel()
        registros = upload_proveedores_to_postgres(df)
        elapsed = time.time() - start_time
        log("=" * 70)
        log(f"  PROVEEDORES COMPLETADO - {registros} registros en {elapsed:.2f}s")
        log("=" * 70)
        return registros
    except Exception as e:
        log(f"ERROR en carga de proveedores: {e}", "error")
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
                    return upload_catalogo_to_postgres(df_catalogo)

                future_ot = executor.submit(ot_upload_task)
                future_cat = executor.submit(cat_upload_task)

                total_ot, counts_ot = future_ot.result()
                total_cat = future_cat.result()
        else:
            df_merged, counts_ot = merge_dataframes(df_primary, df_secondary_mapped, df_memofichas_mapped)
            upload_to_postgres(df_merged)
            total_ot = df_merged.height
            total_cat = None

        elapsed = time.time() - start_time
        log("=" * 70)
        log("  COMPLETADO EXITOSAMENTE")
        log(f"  - basegeneralcostos: {total_ot} registros totales")
        log(f"    - BASE_GENERAL: {counts_ot['BASE_GENERAL']} | POSTVENTA: {counts_ot['POSTVENTA']} | MEMOFICHAS: {counts_ot['MEMOFICHAS']}")
        if total_cat is not None:
            log(f"  - Catalogo: {total_cat} registros")
        log(f"  Tiempo total: {elapsed:.2f} segundos")
        log("=" * 70)

    except Exception as e:
        log(f"ERROR FATAL: {e}", "error")
        raise


if __name__ == "__main__":
    import sys
    import os
    from logging.handlers import RotatingFileHandler
    
    # Configurar logging standalone (sin api_server) - TODO VA A ARCHIVO
    basedir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(basedir, "sync.log")
    
    _logger = logging.getLogger("SyncOT")
    _logger.setLevel(logging.INFO)
    
    formatter = logging.Formatter('[%(asctime)s] %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    file_handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8')
    file_handler.setFormatter(formatter)
    _logger.addHandler(file_handler)
    
    try:
        # Script standalone: carga completa (OT + catalogo) como antes
        upload_buffer_with_merge(include_catalogo=True)
        sys.exit(0)
    except Exception as e:
        log(f"ERROR FATAL: {e}", "error")
        sys.exit(1)
