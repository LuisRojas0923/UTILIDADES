SELECT 
    l.fechaaplicacion::DATE AS "FECHA ENTREGA REPORTE", -- TODO: Reemplazar por ln.fecharealgasto cuando se corrija tabla OT        
    UPPER(l.nombreempleado) AS "NOMBRE", -- TODO: Reemplazar por ln.nombreempleado cuando se corrija tabla OT
    l.empleado::BIGINT AS "DOCUMENTO DE IDENTIDAD", -- TODO: Reemplazar por ln.empleado cuando se corrija tabla OT
    CASE 
        WHEN ln.ot IS NOT NULL AND TRIM(ln.ot) <> '' THEN ln.ot 
        ELSE 'C' || ln.centrocosto 
    END AS "OT-CC", -- TODO: Reemplazar por ln.ot cuando se corrija tabla OT
    COALESCE(ln.fecharealgasto, l.fechaaplicacion)::DATE AS "FECHA REAL DEL GASTO",     -- TODO: Reemplazar por ln.fecharealgasto cuando se corrija tabla OT
    EXTRACT(YEAR FROM COALESCE(ln.fecharealgasto, l.fechaaplicacion))::INTEGER AS "AÑO", -- TODO: Reemplazar por ln.fecharealgasto cuando se corrija tabla OT
    TO_CHAR(COALESCE(ln.fecharealgasto, l.fechaaplicacion), 'TMMonth') AS "MES", -- TODO: Reemplazar por ln.fecharealgasto cuando se corrija tabla OT
    EXTRACT(WEEK FROM COALESCE(ln.fecharealgasto, l.fechaaplicacion))::INTEGER AS "SEMANA DEL AÑO", -- TODO: Reemplazar por ln.fecharealgasto cuando se corrija tabla OT
    o.cliente AS "OBRA", -- TODO: Reemplazar por ln.cliente cuando se corrija tabla OT
    o.ciudad AS "CIUDAD", -- TODO: Reemplazar por ln.ciudad cuando se corrija tabla OT
    ln.centrocosto AS "CENTRO DE COSTO",            
    ln.subcentrocosto AS "SUB CENTRO", -- TODO: Reemplazar por ln.subcentrocosto cuando se corrija tabla OT
    ln.categoria AS "DESCRIPCION", -- TODO: Reemplazar por ln.categoria cuando se corrija tabla OT
    COALESCE(ln.valorconfactura, 0)::BIGINT AS "VALOR TOTAL FACTURA", -- TODO: Reemplazar por ln.valorconfactura cuando se corrija tabla OT
    COALESCE(ln.valorsinfactura, 0)::BIGINT AS "VALOR SIN FACTURA", -- TODO: Reemplazar por ln.valorsinfactura cuando se corrija tabla OT
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
