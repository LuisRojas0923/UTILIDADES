import os
import re
import unicodedata

import polars as pl

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

# Conexion a PostgreSQL
DB_URI = "postgresql://postgres:AdminSolid2025@192.168.0.21:5432/solid"

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
else:
    FILE_PRIMARY = r"\\192.168.0.3\Procesos Comunes SGI\Costos\INFORME DE ORDENES\BASE DE DATOS GENERAL.xlsm"
    FILE_SECONDARY = r"\\192.168.0.3\Postventa\MANTENIMIENTO Y SERVICIO POSTVENTA\- GESTION ORDENES DE SERVICIO\CENTRO LOGÍSTICO\MTZ-SPT-02 Informe Gestion Postventa V1.xlsm"
    FILE_CATALOGO = r"\\192.168.0.3\Procesos Comunes SGI\Mejora\Catalogo de Articulos\CATALOGO FINAL.xlsx"
    FILE_MEMOFICHAS = r"\\192.168.0.3\Control Presupuestal\MEMOFICHA\CONSULTAS\CONSULTA MEMOFICHAS v2.xlsx"

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

