-- Tabla de sincronizacion de proveedores (nit, nombre)
-- Alineada con campos clave de la tabla ERP proveedor en solid.
CREATE TABLE IF NOT EXISTS proveedorprincip (
    nit text PRIMARY KEY,
    nombre text NOT NULL
);
