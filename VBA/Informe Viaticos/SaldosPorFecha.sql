/*
 * CONSULTA: Saldos de Viáticos por Empleado a una Fecha Específica
 * ORIGEN: Extraído de Macro_ConsultaSaldosFecha.bas
 * DESCRIPCIÓN: Calcula el saldo acumulado por empleado hasta una fecha de corte.
 *              Muestra una línea por empleado con su saldo final a esa fecha.
 * BASE DE DATOS: solid
 */

-- Configuración de parámetros para ejecución manual
-- Cambie este valor para filtrar los resultados
WITH params AS (
  SELECT '2025-01-31'::date AS v_fecha  -- Fecha de corte
),
transacciones_hasta_fecha AS (
  SELECT 
    t.empleado,
    t.fechaaplicacion::timestamp,
    t.codigo,
    CASE 
      WHEN t.tipodocumento = 'CONSIGNACION' THEN t.valor::numeric
      WHEN t.tipodocumento = 'LEGALIZACION' THEN -(t.valor::numeric)
      ELSE 0
    END AS movimiento,
    COALESCE(c.nombreempleado, l.nombreempleado) AS nombreempleado
  FROM transaccionviaticos t
  CROSS JOIN params p
  LEFT JOIN consignacion c
    ON c.codigoconsignacion = t.numerodocumento
  LEFT JOIN legalizacion l
    ON l.codigolegalizacion = t.numerodocumento
  WHERE t.fechaaplicacion::timestamp <= (p.v_fecha::timestamp + INTERVAL '1 day' - INTERVAL '1 second')
)
SELECT 
  t.empleado AS "CEDULA",
  MAX(t.nombreempleado) AS "EMPLEADO",
  SUM(t.movimiento) AS "SALDO"
FROM transacciones_hasta_fecha t
GROUP BY t.empleado
ORDER BY t.empleado ASC;
