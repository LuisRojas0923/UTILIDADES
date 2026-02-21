WITH params AS (
    SELECT 
        '2024-01-01'::DATE AS fecha_desde,
        '2025-12-31'::DATE AS fecha_hasta,
        NULL::TEXT AS filtro_empleado -- Cambiar NULL por ID (ej: '16161777')
)
SELECT 
    "RADICADO",
    "DOCUMENTO DE IDENTIDAD",
    "OT-CC",
    "APROBADO",
    "RADICADO" || '-' || ROW_NUMBER() OVER(PARTITION BY "RADICADO" ORDER BY "OT-CC") AS "LLAVE"
FROM (
    -- DETALLE DE GASTOS POR RADICADO Y OT/CENTRO DE COSTO
    SELECT 
        l.codigolegalizacion AS "RADICADO",
        l.empleado AS "DOCUMENTO DE IDENTIDAD",
        CASE 
            WHEN TRIM(COALESCE(ln.ot, '')) = '' THEN 'C' || COALESCE(ln.centrocosto, '')
            ELSE ln.ot 
        END AS "OT-CC",
        SUM(COALESCE(ln.valorsinfactura, 0) + COALESCE(ln.valorconfactura, 0))::BIGINT AS "APROBADO",
        l.fechaaplicacion -- Para el ordenamiento cronológico interno
    FROM 
        linealegalizacion ln
    JOIN 
        legalizacion l ON ln.legalizacion = l.codigo
    CROSS JOIN
        params p
    WHERE 
        (p.fecha_desde IS NULL OR l.fechaaplicacion >= p.fecha_desde)
        AND (p.fecha_hasta IS NULL OR l.fechaaplicacion <= p.fecha_hasta)
        AND (p.filtro_empleado IS NULL OR l.empleado = p.filtro_empleado)
    GROUP BY 
        l.codigolegalizacion, 
        l.empleado,
        l.fechaaplicacion,
        CASE 
            WHEN TRIM(COALESCE(ln.ot, '')) = '' THEN 'C' || COALESCE(ln.centrocosto, '')
            ELSE ln.ot 
        END
) AS grouped_results
ORDER BY fechaaplicacion, "RADICADO", "LLAVE";
