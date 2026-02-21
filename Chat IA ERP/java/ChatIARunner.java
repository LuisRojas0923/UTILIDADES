import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.File;
import java.io.IOException;

/**
 * ChatIARunner - Orquestador Java para el modulo de Chat IA del ERP SOLID.
 * Esta clase ejecuta un script Python que utiliza Google Gemini para generar
 * consultas SQL dinamicas y analizar resultados de la base de datos.
 */
public class ChatIARunner {

    // CONFIGURACION DE RUTAS
    // 1. Ruta sugerida para el servidor de produccion
    private static final String PROD_BASE_DIR = "C:\\ERP\\Chat IA ERP";
    
    // 2. Ruta de ejecucion actual (para pruebas locales sin configurar nada)
    private static final String LOCAL_BASE_DIR = System.getProperty("user.dir");

    public static void main(String[] args) {
        System.out.println("======================================================================");
        System.out.println("--- ERP SOLID: Chat IA - Consultas Inteligentes (Java-Python) ---");
        System.out.println("======================================================================\n");
        
        String userQuery = null;
        
        // Si hay argumentos, usar el primero como consulta
        if (args.length > 0 && !args[0].trim().isEmpty()) {
            userQuery = args[0];
        } else {
            // Si no hay argumentos, leer desde consola
            userQuery = leerConsultaDesdeConsola();
            if (userQuery == null || userQuery.trim().isEmpty()) {
                System.err.println("ERROR: No se proporciono ninguna consulta");
                System.exit(1);
            }
        }
        
        long startTime = System.currentTimeMillis();
        
        int exitCode = runChatIAProcess(userQuery);
        
        long elapsed = System.currentTimeMillis() - startTime;
        
        System.out.println("\n======================================================================");
        System.out.println(String.format("Tiempo Total (Orquestacion Java): %.2f segundos", elapsed / 1000.0));
        if (exitCode == 0) {
            System.out.println("Codigo de Salida: " + exitCode + " (EXITO)");
        } else if (exitCode == 1) {
            System.out.println("Codigo de Salida: " + exitCode + " (ERROR)");
            System.err.println("\n[ERROR] El modelo de IA no esta disponible en este momento.");
            System.err.println("Todos los modelos de Gemini han agotado su cuota o no estan disponibles.");
        } else {
            System.out.println("Codigo de Salida: " + exitCode + " (ERROR)");
        }
        System.out.println("======================================================================");
        
        System.exit(exitCode);
    }

    private static String leerConsultaDesdeConsola() {
        try {
            System.out.print("Ingrese su consulta: ");
            System.out.flush();
            
            BufferedReader reader = new BufferedReader(new InputStreamReader(System.in, "UTF-8"));
            String consulta = reader.readLine();
            
            if (consulta != null) {
                consulta = consulta.trim();
            }
            
            return consulta;
        } catch (IOException e) {
            System.err.println("ERROR al leer desde consola: " + e.getMessage());
            return null;
        }
    }

    private static String getFinalBasePath() {
        // Primero verificar ruta de produccion
        if (new File(PROD_BASE_DIR).exists()) {
            return PROD_BASE_DIR;
        }
        
        // Buscar la carpeta "Chat IA ERP" desde la ubicacion actual
        File currentDir = new File(System.getProperty("user.dir"));
        
        // Buscar hacia arriba en el arbol de directorios
        File searchDir = currentDir;
        while (searchDir != null) {
            String dirName = searchDir.getName();
            
            // Si encontramos "Chat IA ERP", esa es la raiz
            if (dirName.equals("Chat IA ERP")) {
                return searchDir.getAbsolutePath();
            }
            
            // Verificar si dentro de este directorio existe "Chat IA ERP"
            File chatIADir = new File(searchDir, "Chat IA ERP");
            File pythonDir = new File(chatIADir, "python");
            File pythonVenv = new File(pythonDir, ".venv");
            if (pythonVenv.exists()) {
                return chatIADir.getAbsolutePath();
            }
            
            // Si estamos dentro de "Chat IA ERP", verificar si existe la carpeta python
            File pythonDirDirect = new File(searchDir, "python");
            File pythonVenvDirect = new File(pythonDirDirect, ".venv");
            if (pythonVenvDirect.exists()) {
                return searchDir.getAbsolutePath();
            }
            
            // Subir un nivel
            searchDir = searchDir.getParentFile();
        }
        
        // Si no se encuentra, usar directorio actual
        return LOCAL_BASE_DIR;
    }

    private static int runChatIAProcess(String userQuery) {
        try {
            String basePath = getFinalBasePath();
            
            String pythonCmd = basePath + "\\python\\.venv\\Scripts\\python.exe";
            String pythonScript = basePath + "\\python\\chat_ia_erp.py";

            File pythonExe = new File(pythonCmd);
            File scriptFile = new File(pythonScript);

            System.out.println("[INFO] Buscando entorno en: " + basePath);
            System.out.println("[INFO] Consulta del usuario: " + userQuery);

            if (!pythonExe.exists()) {
                System.err.println("ERROR: No se encuentra el entorno virtual en: " + pythonCmd);
                System.err.println("Asegurese de copiar la carpeta 'python' en: " + basePath);
                System.err.println("Y ejecutar 'instalar_entorno.ps1' dentro de ella.");
                return 1;
            }

            if (!scriptFile.exists()) {
                System.err.println("ERROR: No se encuentra el script Python en: " + pythonScript);
                return 1;
            }

            // Construir comando con la consulta como argumento
            ProcessBuilder pb = new ProcessBuilder(pythonCmd, pythonScript, userQuery);
            pb.redirectErrorStream(true);
            
            Process process = pb.start();
            
            // Capturar respuesta de Python (stdout)
            StringBuilder response = new StringBuilder();
            try (BufferedReader reader = new BufferedReader(
                    new InputStreamReader(process.getInputStream(), "UTF-8"))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    // Filtrar mensajes de debug que van a stderr
                    if (line.startsWith("[INFO]") || line.startsWith("[ERROR]")) {
                        System.err.println("[PY] " + line);
                    } else {
                        // La respuesta real va a stdout
                        response.append(line).append("\n");
                    }
                }
            }
            
            int exitCode = process.waitFor();
            
            // Si fue exitoso, imprimir la respuesta (sin los mensajes de debug)
            if (exitCode == 0 && response.length() > 0) {
                System.out.print(response.toString());
            }
            
            return exitCode;
            
        } catch (Exception e) {
            System.err.println("ERROR CRITICO en la orquestacion: " + e.getMessage());
            e.printStackTrace();
            return 1;
        }
    }
}

