/**
 * Clase temporal para probar el Reporte de Gastos de Viáticos.
 */
public class PruebaReporte {
    public static void main(String[] args) {
        try {
            // 1. Iniciar la conexión centralizada del ERP
            ConexionDatos conexion = new ConexionDatos();

            if (conexion.estaConectado()) {
                // 2. Instanciar el reporte
                ReporteGastosViaticos reporte = new ReporteGastosViaticos(conexion);

                // 3. Generar el PDF (usa un código de legalización real de tu BD) RCC-L1056"
                // RCE-L1188
                String codigo = "RCE-L1188";
                System.out.println("Generando reporte para: " + codigo);

                reporte.generar(codigo);

                System.out.println("Proceso finalizado. El PDF debería abrirse automáticamente.");

                conexion.cerrarConexion();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
