import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.io.File;

/**
 * Utilidad Java 8 para consumir el endpoint de exportación a Excel del ERP.
 */
public class EstadoCuentaConsumer {

    private static final String BASE_URL = "http://192.168.40.148:8000/api/v2/viaticos/estado-cuenta/xlsx";
    private static final String DOWNLOAD_PATH = "temp/"; // Ruta relativa al root del proyecto

    /**
     * Descarga el estado de cuenta en formato Excel.
     * 
     * @param cedula Cédula del empleado.
     * @throws Exception Si ocurre un error en la conexión o escritura.
     */
    public void descargarExcel(String cedula) throws Exception {
        String urlStr = BASE_URL + "?cedula=" + cedula;
        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setConnectTimeout(5000);
        conn.setReadTimeout(30000);

        System.out.println("Solicitando reporte para cédula: " + cedula + "...");
        System.out.println("URL: " + urlStr);
        
        int responseCode = conn.getResponseCode();
        if (responseCode == 200) {
            // Asegurar que la carpeta temp existe
            File directory = new File(DOWNLOAD_PATH);
            if (!directory.exists()) {
                directory.mkdirs();
            }

            String fileName = DOWNLOAD_PATH + "Reporte_" + cedula + ".xlsx";
            
            try (InputStream in = conn.getInputStream();
                 OutputStream out = new FileOutputStream(fileName)) {
                
                byte[] buf = new byte[4096];
                int n;
                long totalRead = 0;
                
                while ((n = in.read(buf)) != -1) {
                    out.write(buf, 0, n);
                    totalRead += n;
                }
                
                System.out.println("✓ Excel descargado con éxito: " + fileName);
                System.out.println("  Tamaño: " + (totalRead / 1024) + " KB");
            }
        } else {
            System.err.println("✗ Error al descargar: Código HTTP " + responseCode);
            try (InputStream err = conn.getErrorStream()) {
                if (err != null) {
                    byte[] buf = new byte[4096];
                    int n = err.read(buf);
                    if (n > 0) {
                        System.err.println("  Detalle error: " + new String(buf, 0, n));
                    }
                }
            }
        }
        
        conn.disconnect();
    }

    /**
     * Prueba rápida de ejecución.
     */
    public static void main(String[] args) {
        EstadoCuentaConsumer consumer = new EstadoCuentaConsumer();
        try {
            // Usamos la cédula de ejemplo de la documentación o una de prueba
            String cedulaPrueba = "94041597"; 
            consumer.descargarExcel(cedulaPrueba);
        } catch (Exception e) {
            System.err.println("Ocurrió un error inesperado:");
            e.printStackTrace();
        }
    }
}
