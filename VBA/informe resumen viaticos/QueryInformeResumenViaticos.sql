-- Resumen de viáticos por línea de legalización (PostgreSQL).
-- Optimización: fecha base y montos se calculan una vez por fila (menos CPU).
-- Para más velocidad en tablas grandes, crear índices (ajustar nombres si difieren):
--   CREATE INDEX IF NOT EXISTS ix_linealegalizacion_legalizacion ON linealegalizacion (legalizacion);
--   CREATE INDEX IF NOT EXISTS ix_otviaticos_numero ON otviaticos (numero);
-- Reducir filas con WHERE (fechas, radicado, empleado) suele ayudar más que micro-optimizar el SELECT.

SELECT 
    b.fechaaplicacion_entrega::DATE AS "FECHA ENTREGA REPORTE",
    b.nombre_upper AS "NOMBRE",
    b.empleado::BIGINT AS "DOCUMENTO DE IDENTIDAD",
    b.ot_cc AS "OT-CC",
    b.fe_real AS "FECHA REAL DEL GASTO",
    EXTRACT(YEAR FROM b.fe_real)::INTEGER AS "Año",
    CASE EXTRACT(MONTH FROM b.fe_real)::INTEGER
        WHEN 1 THEN 'enero'
        WHEN 2 THEN 'febrero'
        WHEN 3 THEN 'marzo'
        WHEN 4 THEN 'abril'
        WHEN 5 THEN 'mayo'
        WHEN 6 THEN 'junio'
        WHEN 7 THEN 'julio'
        WHEN 8 THEN 'agosto'
        WHEN 9 THEN 'septiembre'
        WHEN 10 THEN 'octubre'
        WHEN 11 THEN 'noviembre'
        WHEN 12 THEN 'diciembre'
    END AS "Mes",
    EXTRACT(WEEK FROM b.fe_real)::INTEGER AS "SEMANA DEL AÑO",
    b.obra AS "OBRA",
    b.ciudad AS "CIUDAD",
    b.centrocosto AS "CENTRO DE COSTO",
    b.subcentrocosto AS "SUB CENTRO",
    b.categoria AS "DESCRIPCION",
    b.v_conf AS "VALOR TOTAL FACTURA",
    b.v_sin AS "VALOR SIN FACTURA",
    (b.v_conf + b.v_sin) AS "APROBADO",
    (b.v_conf + b.v_sin) AS "SOLICITADO",
    b.codigolegalizacion AS "RADICADO",
    SPLIT_PART(b.codigolegalizacion, '-', 1) AS "AREA",
    0::BIGINT AS "DIFERENCIA"
FROM (
    SELECT 
        l.fechaaplicacion AS fechaaplicacion_entrega,
        UPPER(l.nombreempleado) AS nombre_upper,
        l.empleado,
        l.codigolegalizacion,
        CASE 
            WHEN ln.ot IS NOT NULL AND TRIM(ln.ot) <> '' THEN ln.ot 
            ELSE 'C' || ln.centrocosto 
        END AS ot_cc,
        COALESCE(ln.fecharealgasto, l.fechaaplicacion)::DATE AS fe_real,
        o.cliente AS obra,
        o.ciudad,
        ln.centrocosto,
        ln.subcentrocosto,
        ln.categoria,
        COALESCE(ln.valorconfactura, 0)::BIGINT AS v_conf,
        COALESCE(ln.valorsinfactura, 0)::BIGINT AS v_sin
    FROM legalizacion l
    INNER JOIN linealegalizacion ln ON l.codigo = ln.legalizacion
    LEFT JOIN otviaticos o ON ln.ot = o.numero
) b
ORDER BY b.fechaaplicacion_entrega DESC;
