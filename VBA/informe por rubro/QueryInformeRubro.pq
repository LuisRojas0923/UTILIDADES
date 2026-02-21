SELECT 
    l.fechaaplicacion::DATE AS "FECHA ENTREGA REPORTE",
    UPPER(l.nombreempleado) AS "NOMBRE",
    l.empleado::BIGINT AS "DOCUMENTO DE IDENTIDAD",
    CASE 
        WHEN ln.ot IS NOT NULL AND TRIM(ln.ot) <> '' THEN ln.ot 
        ELSE 'C' || ln.centrocosto 
    END AS "OT-CC",
    COALESCE(ln.fecharealgasto, l.fechaaplicacion)::DATE AS "FECHA REAL DEL GASTO",
    EXTRACT(YEAR FROM COALESCE(ln.fecharealgasto, l.fechaaplicacion))::INTEGER AS "AÑO",
    TO_CHAR(COALESCE(ln.fecharealgasto, l.fechaaplicacion), 'TMMonth') AS "MES",
    EXTRACT(WEEK FROM COALESCE(ln.fecharealgasto, l.fechaaplicacion))::INTEGER AS "SEMANA DEL AÑO",
    o.cliente AS "OBRA",
    o.ciudad AS "CIUDAD",
    ln.centrocosto AS "CENTRO DE COSTO",
    ln.subcentrocosto AS "SUB CENTRO",
    ln.categoria AS "DESCRIPCION",
    COALESCE(ln.valorconfactura, 0)::BIGINT AS "VALOR TOTAL FACTURA",
    COALESCE(ln.valorsinfactura, 0)::BIGINT AS "VALOR SIN FACTURA",
    (COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT AS "APROBADO",
    (COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT AS "SOLICITADO",
    l.codigolegalizacion AS "RADICADO",
    SPLIT_PART(l.codigolegalizacion, '-', 1) AS "AREA",
    0::BIGINT AS "DIFERENCIA"
FROM 
    legalizacion l
JOIN 
    linealegalizacion ln ON l.codigo = ln.legalizacion
LEFT JOIN otviaticos o ON ln.ot = o.numero
ORDER BY 
    l.fechaaplicacion DESC;
