SELECT
    l.empleado::BIGINT AS "DOCUMENTO DE IDENTIDAD",
    ln.ot AS "OT-CC",
    ln.centrocosto AS "CENTRO DE COSTO",
    ln.subcentrocosto AS "SUB CENTRO",
    l.codigolegalizacion AS "RADICADO",
    SUM(COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT AS "VALOR OT"
FROM
    linealegalizacion ln
JOIN
    legalizacion l ON ln.legalizacion = l.codigo
WHERE
    ln.ot IS NOT NULL AND TRIM(ln.ot) <> ''
    AND UPPER(l.estado) = 'CONTABILIZADO'
GROUP BY
    l.empleado,
    ln.ot,
    ln.centrocosto,
    ln.subcentrocosto,
    l.codigolegalizacion
ORDER BY
    l.codigolegalizacion ASC, ln.ot ASC;
