/**
 * Clase para el botón de 'Imprimir Legalización' en el ERP.
 */
public class ImprimirLegalizacion {
    public static void main(String[] args) {
        try {
            // 1. Iniciar la conexión centralizada del ERP
            ConexionDatos conexion = new ConexionDatos();

            // Método oficial de conexión en el ERP
            int status = conexion.conexionPostgreSQL("192.168.0.21:5432", "solid", "postgres", "AdminSolid2025");

            if (status == 1) {
                // 2. Instanciar el reporte
                ReporteGastosViaticos reporte = new ReporteGastosViaticos(conexion);

                // 3. Obtener el código
                String codigo = "ADN-L1038";
                if (args.length > 0) {
                    codigo = args[0];
                }

                System.out.println("Generando reporte para: " + codigo);
                reporte.generar(codigo);

                System.out.println("Proceso finalizado. El PDF deberia abrirse automaticamente.");
                conexion.terminarConexion();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
