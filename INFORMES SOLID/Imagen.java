/**
 * Clase stub para Imagen
 */
public class Imagen {
    public Object representacion = null;
    
    public ConexionDatos conexion;
    
    public Imagen(Object obj, ConexionDatos conexion) {
        this.conexion = conexion;
    }
    
    public void cargar(String campo, String valor) {
        // Cargar imagen desde BD
        try {
            String sql = "SELECT representacion FROM imagen WHERE " + campo + " = '" + valor + "'";
            java.sql.ResultSet rs = conexion.funcionConsultar(sql);
            if (rs != null && rs.next()) {
                representacion = rs.getObject("representacion");
            }
        } catch (Exception e) {
            // Si falla, representacion queda null
        }
    }
}
