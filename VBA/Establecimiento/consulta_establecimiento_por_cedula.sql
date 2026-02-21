-- Consulta por cédula: establecimiento + contrato (cargo, area, estado, ciudadcontratacion, centrocosto)
-- Misma convención que consulta_establecimiento.sql (E, C; TRIM/CAST en JOINs).
-- En pgAdmin reemplazar :cedula por el valor, ej: '1146437946' o 1146437946 según tipo de columna.
-- Estado activo: C.estado = 'Activo' (en BD); si se usa 'A' en app, filtrar por C.estado = 'Activo'.

SELECT DISTINCT ON (E.nrocedula)
    E.nrocedula      AS "nrocedula",
    E.nombre::text   AS "nombre",
    C.cargo::text    AS "cargo",
    C.area::text     AS "area",
    C.estado::text   AS "estado",
    C.ciudadcontratacion::text AS "ciudadcontratacion",
    E.viaticante,
    E.baseviaticos,
    C.centrocosto::text AS "centrocosto"
FROM establecimiento E
LEFT JOIN contrato C
    ON TRIM(CAST(C.establecimiento AS TEXT)) = TRIM(CAST(E.nrocedula AS TEXT))
WHERE TRIM(CAST(E.nrocedula AS TEXT)) = TRIM(CAST(:cedula AS TEXT))
  AND C.estado = 'Activo'
ORDER BY E.nrocedula, C.fechainicio DESC NULLS LAST;

-- Para pgAdmin: usar la consulta de abajo reemplazando '1146437946' por la cédula deseada.
/*
SELECT DISTINCT ON (E.nrocedula)
    E.nrocedula      AS "nrocedula",
    E.nombre::text   AS "nombre",
    C.cargo::text    AS "cargo",
    C.area::text     AS "area",
    C.estado::text   AS "estado",
    C.ciudadcontratacion::text AS "ciudadcontratacion",
    E.viaticante,
    E.baseviaticos,
    C.centrocosto::text AS "centrocosto"
FROM establecimiento E
LEFT JOIN contrato C
    ON TRIM(CAST(C.establecimiento AS TEXT)) = TRIM(CAST(E.nrocedula AS TEXT))
WHERE TRIM(CAST(E.nrocedula AS TEXT)) = '1146437946'
  AND C.estado = 'Activo'
ORDER BY E.nrocedula, C.fechainicio DESC NULLS LAST;
*/
