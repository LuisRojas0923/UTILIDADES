import java.sql.ResultSet;
import java.sql.ResultSetMetaData;

public class SchemaInspector {
    public static void main(String[] args) {
        try {
            ConexionDatos conexion = new ConexionDatos();
            int status = conexion.conexionPostgreSQL("192.168.0.21:5432", "solid", "postgres", "AdminSolid2025");

            if (status == 1) {
                String sql = "SELECT * FROM linealegalizacion LIMIT 1";
                ResultSet rs = conexion.funcionConsultar(sql);
                if (rs != null) {
                    ResultSetMetaData meta = rs.getMetaData();
                    int count = meta.getColumnCount();
                    System.out.println("--- Columnas de linealegalizacion ---");
                    for (int i = 1; i <= count; i++) {
                        // El índice en datosL será i-1
                        System.out.println("Index " + (i - 1) + ": " + meta.getColumnName(i));
                    }
                }
                conexion.terminarConexion();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
