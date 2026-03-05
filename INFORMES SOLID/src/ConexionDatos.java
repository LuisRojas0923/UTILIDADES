/*
 * Clase de conexión a base de datos PostgreSQL
 * Para pruebas y desarrollo
 * 
 * Parámetros de conexión:
 * - Servidor: 192.168.0.21
 * - Puerto: 5432
 * - Base de datos: solid
 * - Usuario: postgres
 * - Contraseña: AdminSolid2025
 */

import java.sql.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Clase para manejar la conexión a la base de datos PostgreSQL
 */
public class ConexionDatos {
    
    // Parámetros de conexión
    private static final String SERVIDOR = "192.168.0.21";
    private static final String PUERTO = "5432";
    private static final String BASE_DATOS = "solid";
    private static final String USUARIO = "postgres";
    private static final String CONTRASENA = "AdminSolid2025";
    
    // URL de conexión JDBC
    private static final String URL = "jdbc:postgresql://" + SERVIDOR + ":" + PUERTO + "/" + BASE_DATOS;
    
    // Conexión activa
    private Connection conexion = null;
    
    /**
     * Constructor - establece la conexión automáticamente
     */
    public ConexionDatos() {
        conectar();
    }
    
    /**
     * Establece la conexión a la base de datos
     */
    private void conectar() {
        try {
            // Cargar el driver de PostgreSQL
            Class.forName("org.postgresql.Driver");
            
            // Establecer conexión
            conexion = DriverManager.getConnection(URL, USUARIO, CONTRASENA);
            
            if (conexion != null) {
                System.out.println("✓ Conexión establecida exitosamente a PostgreSQL");
                System.out.println("  Servidor: " + SERVIDOR);
                System.out.println("  Base de datos: " + BASE_DATOS);
            }
            
        } catch (ClassNotFoundException e) {
            System.err.println("Error: Driver de PostgreSQL no encontrado");
            System.err.println("Asegúrate de tener el JAR de PostgreSQL en el classpath");
            e.printStackTrace();
        } catch (SQLException e) {
            System.err.println("Error al conectar a la base de datos:");
            System.err.println("  URL: " + URL);
            e.printStackTrace();
        }
    }
    
    /**
     * Ejecuta una consulta SQL y retorna un ResultSet
     * 
     * @param sql - Consulta SQL a ejecutar
     * @return ResultSet con los resultados
     */
    public ResultSet funcionConsultar(String sql) {
        try {
            if (conexion == null || conexion.isClosed()) {
                System.out.println("Reconectando a la base de datos...");
                conectar();
            }
            
            if (conexion != null) {
                Statement stmt = conexion.createStatement();
                ResultSet rs = stmt.executeQuery(sql);
                return rs;
            }
            
        } catch (SQLException e) {
            System.err.println("Error al ejecutar consulta SQL:");
            System.err.println("SQL: " + sql);
            e.printStackTrace();
        }
        
        return null;
    }
    
    /**
     * Convierte un ResultSet en un arreglo bidimensional de Strings
     * 
     * @param rs - ResultSet a convertir
     * @return Arreglo bidimensional con los datos
     */
    public static String[][] armarArreglo(ResultSet rs) {
        if (rs == null) {
            return new String[][]{{"0"}};
        }
        
        List<String[]> filas = new ArrayList<String[]>();
        
        try {
            // Obtener información de las columnas
            ResultSetMetaData metaData = rs.getMetaData();
            int numColumnas = metaData.getColumnCount();
            
            // Procesar cada fila
            while (rs.next()) {
                String[] fila = new String[numColumnas];
                for (int i = 1; i <= numColumnas; i++) {
                    Object valor = rs.getObject(i);
                    fila[i - 1] = (valor != null) ? valor.toString() : "";
                }
                filas.add(fila);
            }
            
            // Convertir lista a arreglo bidimensional
            if (filas.isEmpty()) {
                return new String[][]{{"0"}};
            }
            
            String[][] resultado = new String[filas.size()][];
            for (int i = 0; i < filas.size(); i++) {
                resultado[i] = filas.get(i);
            }
            
            return resultado;
            
        } catch (SQLException e) {
            System.err.println("Error al procesar ResultSet:");
            e.printStackTrace();
            return new String[][]{{"0"}};
        }
    }
    
    /**
     * Cierra la conexión a la base de datos
     */
    public void cerrarConexion() {
        try {
            if (conexion != null && !conexion.isClosed()) {
                conexion.close();
                System.out.println("Conexión cerrada");
            }
        } catch (SQLException e) {
            System.err.println("Error al cerrar la conexión:");
            e.printStackTrace();
        }
    }
    
    /**
     * Verifica si la conexión está activa
     * 
     * @return true si está conectado, false en caso contrario
     */
    public boolean estaConectado() {
        try {
            return conexion != null && !conexion.isClosed();
        } catch (SQLException e) {
            return false;
        }
    }
    
    /**
     * Obtiene la conexión directa (para uso avanzado)
     * 
     * @return Connection object
     */
    public Connection getConexion() {
        return conexion;
    }
    
    /**
     * Método de prueba para verificar la conexión
     */
    public static void main(String[] args) {
        System.out.println("=== Prueba de Conexión a PostgreSQL ===\n");
        
        ConexionDatos conexion = new ConexionDatos();
        
        if (conexion.estaConectado()) {
            System.out.println("\n✓ Conexión exitosa!");
            
            // Prueba de consulta simple
            try {
                String sql = "SELECT version();";
                ResultSet rs = conexion.funcionConsultar(sql);
                
                if (rs != null && rs.next()) {
                    System.out.println("\nVersión de PostgreSQL:");
                    System.out.println("  " + rs.getString(1));
                }
                
            } catch (SQLException e) {
                System.err.println("Error en consulta de prueba:");
                e.printStackTrace();
            }
            
            conexion.cerrarConexion();
        } else {
            System.err.println("\n✗ No se pudo establecer la conexión");
        }
    }
}
