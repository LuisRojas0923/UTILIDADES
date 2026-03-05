/**
 * Runner para ejecutar el Reporte de Gastos de Viáticos.
 */
public class RunnerReporte {
    public static void main(String[] args) {
        try {
            System.out.println("=== Iniciando Generación de Reporte ===");

            // 1. Inicializar conexión
            ConexionDatos conexion = new ConexionDatos();

            // Lógica de conexión del ERP SOLID
            int status = conexion.conexionPostgreSQL("192.168.0.21:5432", "solid", "postgres", "AdminSolid2025");

            if (status == 1) {
                // 2. Instanciar el reporte
                ReporteGastosViaticos reporte = new ReporteGastosViaticos(conexion);

                // 3. Ejecutar para un código de legalización específico
                String codigoPrueba = "OP-L1011";
                System.out.println("Generando reporte para: " + codigoPrueba);

                reporte.generar(codigoPrueba);

                System.out.println("✓ Reporte generado exitosamente.");
            } else {
                System.err.println("✗ Error: No se pudo establecer la conexión a la base de datos.");
            }

        } catch (Exception e) {
            System.err.println("✗ Error durante la ejecución:");
            e.printStackTrace();
        }
    }
}
