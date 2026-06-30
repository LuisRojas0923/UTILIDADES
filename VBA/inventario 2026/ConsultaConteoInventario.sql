-- ====================================================================
-- CONTEO FÍSICO DE INVENTARIO (sin duplicados + optimizado)
-- Base de datos: project_manager (Portal)
-- Tablas: conteoinventario + asignacioninventario
-- Cambiar c1 por c2 o c3 según la ronda de conteo
-- FIX: LEFT JOIN LATERAL para evitar duplicados por múltiples asignaciones
-- ====================================================================

-- ====================================================================
-- ÍNDICES RECOMENDADOS (ejecutar una sola vez en la BD)
-- ====================================================================
-- CREATE INDEX IF NOT EXISTS idx_conteo_ubicacion_codigo
--     ON conteoinventario (bodega, bloque, estante, nivel, codigo);
--
-- CREATE INDEX IF NOT EXISTS idx_conteo_user_c1
--     ON conteoinventario (user_c1) WHERE user_c1 IS NOT NULL AND user_c1 <> '';
--
-- CREATE INDEX IF NOT EXISTS idx_conteo_user_c2
--     ON conteoinventario (user_c2) WHERE user_c2 IS NOT NULL AND user_c2 <> '';
--
-- CREATE INDEX IF NOT EXISTS idx_conteo_user_c3
--     ON conteoinventario (user_c3) WHERE user_c3 IS NOT NULL AND user_c3 <> '';
--
-- CREATE INDEX IF NOT EXISTS idx_asignacion_ubicacion
--     ON asignacioninventario (bodega, bloque, estante, nivel, id);
-- ====================================================================

SELECT 
    ROW_NUMBER() OVER (
        ORDER BY c.bodega, c.bloque, c.estante, c.nivel, c.codigo
    ) AS "No.",
    a.id AS "No. Planilla",
    '' AS "Column3",
    c.b_siigo AS "B. Siigo",
    c.bodega AS "Bodega",
    c.bloque AS "Bloque",
    c.estante AS "Estante",
    c.nivel AS "Nivel",
    '' AS "Column9",
    c.codigo AS "Codigo",
    c.descripcion AS "Descripcion",
    c.unidad AS "Und.",
    '' AS "Column13",
    c.cant_c1 AS "Cant.",
    '' AS "Column15",
    COALESCE(c.obs_c1, '') AS "Observaciones:",
    COALESCE(a.numero_pareja::text, '') AS "DIGITADOR"
FROM conteoinventario c
LEFT JOIN LATERAL (
    SELECT ai.id, ai.numero_pareja
    FROM asignacioninventario ai
    WHERE ai.bodega  = c.bodega
      AND ai.bloque  = c.bloque
      AND ai.estante = c.estante
      AND ai.nivel   = c.nivel
    ORDER BY ai.id
    LIMIT 1
) a ON true
WHERE c.user_c1 IS NOT NULL AND c.user_c1 <> ''
ORDER BY c.bodega, c.bloque, c.estante, c.nivel, c.codigo;
