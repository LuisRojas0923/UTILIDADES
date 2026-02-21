/*
 * CONSULTA: Estado de Cuenta de Viaticos
 * ORIGEN: Extraido de Macro_ConsultaPostgreSQL.bas
 * DESCRIPCION: Calcula el saldo acumulado por empleado cruzando consignaciones y legalizaciones.
 * BASE DE DATOS: solid
 */

-- Configuracion de parametros para ejecucion manual
-- Cambie estos valores para filtrar los resultados
WITH params AS (
  SELECT
    '38561178'::text AS v_empleado,         -- Ejemplo: '16162075'
    NULL::text AS v_nombreempleado,   -- Ejemplo: 'gladys'
    '2026-01-01'::date AS v_desde,    -- Fecha inicial
    '2026-12-31'::date AS v_hasta     -- Fecha final
)
SELECT 
    t.codigo AS "CODIGO",
    t.fechaaplicacion::timestamp AS "FECHA APLICACION",
    t.empleado AS "CEDULA",
    COALESCE(c.nombreempleado, l.nombreempleado) AS "EMPLEADO",
    t.numerodocumento AS "RADICADO",
    CASE WHEN t.tipodocumento = 'CONSIGNACION' THEN t.valor ELSE 0 END AS "VALOR CONSIGNACION",
    CASE WHEN t.tipodocumento = 'LEGALIZACION' THEN t.valor ELSE 0 END AS "VALOR LEGALIZACION",
    SUM(
        CASE 
            WHEN t.tipodocumento = 'CONSIGNACION' THEN t.valor
            WHEN t.tipodocumento = 'LEGALIZACION' THEN -(t.valor)
            ELSE 0
        END
    ) OVER (
        PARTITION BY t.empleado
        ORDER BY t.fechaaplicacion::timestamp ASC, t.codigo ASC
    ) AS "SALDO",
    COALESCE(c.observaciones, l.observaciones) AS "OBSERVACIONES"
FROM (
    -- Agrupar transacciones por documento para unificar lineas de detalle
    SELECT
        MIN(codigo) AS codigo,
        fechaaplicacion,
        empleado,
        numerodocumento,
        tipodocumento,
        SUM(valor::numeric) AS valor
    FROM transaccionviaticos
    GROUP BY fechaaplicacion, empleado, numerodocumento, tipodocumento
) t
CROSS JOIN params p
LEFT JOIN (
    -- Una sola fila por consignacion
    SELECT codigoconsignacion, MAX(nombreempleado) AS nombreempleado, MAX(observaciones) AS observaciones
    FROM consignacion
    GROUP BY codigoconsignacion
) c ON c.codigoconsignacion = t.numerodocumento
LEFT JOIN (
    -- Una sola fila por legalizacion
    SELECT codigolegalizacion, MAX(nombreempleado) AS nombreempleado, MAX(observaciones) AS observaciones
    FROM legalizacion
    GROUP BY codigolegalizacion
) l ON l.codigolegalizacion = t.numerodocumento
WHERE 1=1
  AND (p.v_empleado IS NULL OR t.empleado = p.v_empleado)
  AND (
        p.v_nombreempleado IS NULL
        OR COALESCE(c.nombreempleado, l.nombreempleado) ILIKE ('%' || p.v_nombreempleado || '%')
      )
  AND t.fechaaplicacion::timestamp >= COALESCE(p.v_desde::timestamp, '1900-01-01'::timestamp)
  AND t.fechaaplicacion::timestamp < (COALESCE(p.v_hasta::timestamp, '9999-12-31'::timestamp) + INTERVAL '1 day')
ORDER BY t.empleado, t.fechaaplicacion::timestamp ASC, t.codigo ASC;
