SELECT 
    UPPER(l.nombreempleado) AS "EMPLEADO",
    SUM(COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT AS "APROBADO",
    ln.ot AS "OT-CC",
    ln.centrocosto AS "CC",
    ln.subcentrocosto AS "SUB CENTRO",
    l.empleado::BIGINT AS "CEDULA",
    l.codigolegalizacion AS "RADICADO",
    (l.empleado || '-' || ln.ot || '-' || SUM(COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT) AS "LLAVE"
FROM 
    linealegalizacion ln
JOIN 
    legalizacion l ON ln.legalizacion = l.codigo
WHERE 
    ln.ot IS NOT NULL AND TRIM(ln.ot) <> ''
    AND UPPER(l.estado) = 'CONTABILIZADO'
GROUP BY 
    l.nombreempleado,
    ln.ot,
    ln.centrocosto,
    ln.subcentrocosto,
    l.empleado,
    l.codigolegalizacion
ORDER BY 
    l.nombreempleado ASC, l.codigolegalizacion ASC;
