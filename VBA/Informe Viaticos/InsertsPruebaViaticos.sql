/*
 * SCRIPT DE INSERCIÓN PARA PRUEBAS - INFORME DE VIÁTICOS
 * Propósito: Insertar registros en diferentes estados para validar el reporte EstadoDeCuentaV2.sql
 * Empleado de prueba: 94041597 (ALVARO ANDRES TROCCOLI ESCOBAR)
 */

-- 1. Insertar Consignaciones en diferentes estados
INSERT INTO consignacion (
    codigo, codigoconsignacion, fecha, hora, fechaaplicacion, 
    empleado, nombreempleado, area, valor, estado, 
    usuario, observaciones, anexo, centrocosto, cargo, ciudad
) VALUES 
-- Una consignación ya contabilizada (Saldo base)
(9001, 'TEST-C100', CURRENT_DATE, CURRENT_TIME, CURRENT_DATE, 
 '94041597', 'TROCCOLI ESCOBAR ALVARO ANDRES', 'ADN', 5000000, 'CONTABILIZADO', 
 1, 'CONSIGNACION DE PRUEBA - CONTABILIZADA', 0, '3080-99', 'INGENIERO', 'BOGOTA'),

-- Una consignación en firme (Firmada)
(9002, 'TEST-C101', CURRENT_DATE, CURRENT_TIME, CURRENT_DATE, 
 '94041597', 'TROCCOLI ESCOBAR ALVARO ANDRES', 'ADN', 1000000, 'EN FIRME', 
 1, 'CONSIGNACION DE PRUEBA - EN FIRME (FIRMADA)', 0, '3080-99', 'INGENIERO', 'BOGOTA'),

-- Una consignación pendiente
(9003, 'TEST-C102', CURRENT_DATE, CURRENT_TIME, CURRENT_DATE, 
 '94041597', 'TROCCOLI ESCOBAR ALVARO ANDRES', 'ADN', 500000, 'PENDIENTE', 
 1, 'CONSIGNACION DE PRUEBA - PENDIENTE', 0, '3080-99', 'INGENIERO', 'BOGOTA');


-- 2. Insertar Legalizaciones en diferentes estados
INSERT INTO legalizacion (
    codigo, codigolegalizacion, fecha, hora, fechaaplicacion, 
    empleado, nombreempleado, area, valortotal, estado, 
    usuario, observaciones, anexo, centrocosto, cargo, ciudad
) VALUES 
-- Una legalización en firme (Gasto ya firmado)
(9004, 'TEST-L100', CURRENT_DATE, CURRENT_TIME, CURRENT_DATE, 
 '94041597', 'TROCCOLI ESCOBAR ALVARO ANDRES', 'ADN', 800000, 'EN FIRME', 
 1, 'LEGALIZACION DE PRUEBA - EN FIRME', 0, '3080-99', 'INGENIERO', 'BOGOTA'),

-- Una legalización pendiente (Gasto reportado pero no aprobado)
(9005, 'TEST-L101', CURRENT_DATE, CURRENT_TIME, CURRENT_DATE, 
 '94041597', 'TROCCOLI ESCOBAR ALVARO ANDRES', 'ADN', 300000, 'PENDIENTE', 
 1, 'LEGALIZACION DE PRUEBA - PENDIENTE', 0, '3080-99', 'INGENIERO', 'BOGOTA');

-- NOTA: Después de ejecutar estos inserts, puede correr EstadoDeCuentaV2.sql 
-- filtrando por el empleado '94041597' para ver cómo se desglosan los valores.
