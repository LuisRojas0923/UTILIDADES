-- Áreas únicas (prefijo del radicado antes del guion).
-- Ejemplo codigolegalizacion: OP-5286 -> AREA = OP

SELECT DISTINCT
    UPPER(TRIM(SPLIT_PART(l.codigolegalizacion, '-', 1))) AS "AREA"
FROM
    legalizacion l
WHERE
    l.codigolegalizacion IS NOT NULL
    AND TRIM(l.codigolegalizacion) <> ''
    AND POSITION('-' IN l.codigolegalizacion) > 0
ORDER BY
    1;
