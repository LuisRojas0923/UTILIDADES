/**
 * Stub de la clase Empresa para permitir la compilación.
 */
public class Empresa {
    public String nombre = "EMPRESA EJEMPLO";
    public String nit = "900.000.000-1";
    public String ciudad = "BOGOTA";
    public String departamento = "CUNDINAMARCA";
    public String pais = "COLOMBIA";
    public String direccion = "Calle 123";
    public String telefono = "1234567";
    public String fax = "1234567";
    public String email = "info@ejemplo.com";
    public String sitioWeb = "www.ejemplo.com";
    public String logo = "";

    public Empresa() {
    }

    public Empresa(ConexionDatos conexion) {
    }

    public Empresa(Object frame, ConexionDatos conexion) {
    }

    public void cargar(String campo, String valor) {
        // No hace nada para el reporte de viáticos
    }
}
