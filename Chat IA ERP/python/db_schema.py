"""
Esquema de base de datos para el módulo de Chat IA ERP.
Proporciona información sobre las tablas disponibles para consultas.
"""

# Esquema de la tabla legalizacion
LEGALIZACION_SCHEMA = """
Tabla: legalizacion
Descripcion: Tabla principal de legalizaciones de gastos
Campos principales:
  - codigo (PK): Código único de la legalización
  - codigolegalizacion: Código de radicado (formato: AREA-NUMERO)
  - empleado: Documento de identidad del empleado
  - nombreempleado: Nombre del empleado
  - fechaaplicacion: Fecha de aplicación/entrega del reporte
Relaciones:
  - Uno a muchos con linealegalizacion (legalizacion.codigo = linealegalizacion.legalizacion)
"""

# Esquema de la tabla linealegalizacion
LINEALEGALIZACION_SCHEMA = """
Tabla: linealegalizacion
Descripcion: Detalle de gastos por cada legalización
Campos principales:
  - legalizacion (FK): Referencia a legalizacion.codigo
  - ot: Número de orden de trabajo (puede estar vacío)
  - centrocosto: Centro de costo (usado si ot está vacío)
  - subcentrocosto: Subcentro de costo
  - categoria: Descripción/categoría del gasto
  - valorconfactura: Valor con factura
  - valorsinfactura: Valor sin factura
  - fecharealgasto: Fecha real del gasto
Reglas de negocio:
  - Si ot está vacío o NULL, se usa 'C' || centrocosto como identificador
  - El valor total aprobado es: valorconfactura + valorsinfactura
Relaciones:
  - Muchos a uno con legalizacion (linealegalizacion.legalizacion = legalizacion.codigo)
  - LEFT JOIN opcional con otviaticos (linealegalizacion.ot = otviaticos.numero)
"""

# Esquema de la tabla consignacion
CONSIGNACION_SCHEMA = """
Tabla: consignacion
Descripcion: Información de consignaciones a empleados
Campos principales:
  - codigoconsignacion: Código único de la consignación (contrato)
  - empleado: Documento de identidad del empleado
  - nombreempleado: Nombre del empleado
  - valor: Valor de la consignación
  - estado: Estado de la consignación (ej: 'CONTABILIZADO')
Relaciones:
  - Se relaciona con transaccionviaticos (consignacion.codigoconsignacion = transaccionviaticos.numerodocumento)
  - WHERE transaccionviaticos.tipodocumento = 'CONSIGNACION'
Reglas de negocio:
  - Impuesto 4x1000: valor * 0.004
  - Total consignación: valor + (valor * 0.004)
"""

# Esquema completo para el LLM
FULL_SCHEMA_CONTEXT = f"""
{LEGALIZACION_SCHEMA}

{LINEALEGALIZACION_SCHEMA}

{CONSIGNACION_SCHEMA}

NOTAS IMPORTANTES:
1. Todas las fechas están en formato DATE de PostgreSQL
2. Los valores monetarios son BIGINT (enteros, sin decimales)
3. Para legalizaciones, el total aprobado es la suma de valorconfactura + valorsinfactura
4. Para consignaciones, el impuesto 4x1000 se calcula como valor * 0.004
5. Las tablas relacionadas opcionales incluyen: otviaticos, transaccionviaticos
"""

# Ejemplos de queries comunes
EXAMPLE_QUERIES = """
Ejemplos de consultas SQL comunes:

1. Contar legalizaciones por mes:
   SELECT 
     EXTRACT(YEAR FROM fechaaplicacion) AS año,
     TO_CHAR(fechaaplicacion, 'TMMonth') AS mes,
     COUNT(*) AS total_legalizaciones
   FROM legalizacion
   GROUP BY año, mes
   ORDER BY año DESC, mes DESC;

2. Gastos por empleado:
   SELECT 
     l.empleado,
     l.nombreempleado,
     SUM(ln.valorconfactura + ln.valorsinfactura) AS total_gastos
   FROM legalizacion l
   JOIN linealegalizacion ln ON l.codigo = ln.legalizacion
   GROUP BY l.empleado, l.nombreempleado
   ORDER BY total_gastos DESC;

3. Consignaciones por empleado:
   SELECT 
     c.empleado,
     c.nombreempleado,
     SUM(c.valor) AS total_consignaciones
   FROM consignacion c
   WHERE UPPER(c.estado) LIKE '%CONTABILIZADO%'
   GROUP BY c.empleado, c.nombreempleado
   ORDER BY total_consignaciones DESC;

4. Legalizaciones con detalle:
   SELECT 
     l.codigolegalizacion AS radicado,
     l.nombreempleado,
     l.fechaaplicacion,
     CASE 
       WHEN ln.ot IS NOT NULL AND TRIM(ln.ot) <> '' THEN ln.ot 
       ELSE 'C' || ln.centrocosto 
     END AS ot_cc,
     ln.categoria,
     (ln.valorconfactura + ln.valorsinfactura) AS valor_aprobado
   FROM legalizacion l
   JOIN linealegalizacion ln ON l.codigo = ln.legalizacion
   ORDER BY l.fechaaplicacion DESC;
"""

def get_schema_for_llm():
    """Retorna el esquema completo formateado para el LLM"""
    return FULL_SCHEMA_CONTEXT + "\n\n" + EXAMPLE_QUERIES

