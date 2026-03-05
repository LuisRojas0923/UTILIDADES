/**
 * Runner para ejecutar el Reporte de Estado de Cuenta.
 */
public class RunnerEstadoCuenta {
    public static void main(String[] args) {
        try {
            System.out.println("=== Iniciando Generación de Estado de Cuenta ===");

            // 1. Inicializar conexión
            ConexionDatos conexion = new ConexionDatos();

            // Parámetros de conexión (según RunnerReporte.java)
            int status = conexion.conexionPostgreSQL("192.168.0.21:5432", "solid", "postgres", "AdminSolid2025");

            if (status == 1) {
                // 2. Instanciar el reporte
                ReporteEstadoCuenta reporte = new ReporteEstadoCuenta(conexion);

                // 3. Ejecutar para una cédula específica
                // Se puede cambiar este valor o recibirlo por consola
                String cedulaPrueba = "1143849288";
                System.out.println("Generando estado de cuenta para la cédula: " + cedulaPrueba);

                reporte.generar(cedulaPrueba);

                System.out.println("✓ Reporte generado y abierto exitosamente.");
            } else {
                System.err.println("✗ Error: No se pudo establecer la conexión a la base de datos.");
            }

        } catch (Exception e) {
            System.err.println("✗ Error durante la ejecución:");
            e.printStackTrace();
        }
    }
}
