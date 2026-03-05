/**
 * Clase para manejar la carga de imágenes desde la base de datos.
 */
public class Imagen {
    public Object representacion = null;
    public ConexionDatos conexion;

    public Imagen(Object obj, ConexionDatos conexion) {
        this.conexion = conexion;
    }

    public void cargar(String campo, String valor) {
        try {
            String sql = "SELECT representacion FROM imagen WHERE " + campo + " = '" + valor + "'";
            java.sql.ResultSet rs = conexion.funcionConsultar(sql);
            if (rs != null && rs.next()) {
                representacion = rs.getObject("representacion");
            }
        } catch (Exception e) {
            // Manejo de error silencioso según diseño original
        }
    }
}
