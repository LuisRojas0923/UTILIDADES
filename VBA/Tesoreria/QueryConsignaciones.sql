SELECT 
    EXTRACT(YEAR FROM t.fechaaplicacion)::INTEGER AS "AÑO",
    TO_CHAR(t.fechaaplicacion, 'TMMonth') AS "MES",
    EXTRACT(DAY FROM t.fechaaplicacion)::INTEGER AS "DIA",
    t.fechaaplicacion::DATE AS "FECHA",
    t.empleado::BIGINT AS "CEDULA",
    UPPER(COALESCE(c.nombreempleado, '')) AS "EMPLEADO",
    t.numerodocumento AS "CONTRATO",
    t.valor::BIGINT AS "CONSIGNACION",
    (t.valor * 0.004)::BIGINT AS "IMP 4 X 1000",
    (t.valor + (t.valor * 0.004))::BIGINT AS "TOTAL CONSIGNACION",
    NULL::TEXT AS "VIATICO A PAGAR?"
FROM transaccionviaticos t
LEFT JOIN consignacion c ON t.numerodocumento = c.codigoconsignacion
WHERE t.tipodocumento = 'CONSIGNACION'
    AND UPPER(c.estado) LIKE '%CONTABILIZADO%'
ORDER BY t.fechaaplicacion DESC;
