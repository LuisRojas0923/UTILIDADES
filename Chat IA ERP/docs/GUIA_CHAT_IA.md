# Guía de Despliegue: Chat IA ERP para SOLID

## 1. Resumen

Este módulo permite realizar consultas inteligentes a la base de datos del ERP SOLID usando Google Gemini 2.5 Flash. El sistema genera SQL dinámicamente basado en consultas en lenguaje natural y analiza los resultados para proporcionar respuestas comprensibles.

## 2. Arquitectura

- **Java (ChatIARunner.java)**: Orquestador que se integra con SOLID
- **Python (chat_ia_erp.py)**: Motor de IA que genera SQL y analiza resultados
- **PostgreSQL**: Base de datos con tablas legalizacion, linealegalizacion, consignacion

## 3. Preparación del Entorno

### 3.1 Instalación de Python y Dependencias

1. Navegar a la carpeta `python/`:
```powershell
cd "Chat IA ERP\python"
```

2. Ejecutar el script de instalación:
```powershell
.\instalar_entorno.ps1
```

Este script:
- Instala `uv` si no está presente
- Crea el entorno virtual `.venv`
- Instala todas las dependencias de `requirements.txt`

### 3.2 Configuración de API Key de Google Gemini

Crear archivo `.env` en la carpeta `python/`:

```env
GOOGLE_AI_API_KEY=AIzaSyAAUESKW2_HA5FVH5zq3X0Lg_VKEN56v6M
```

**IMPORTANTE**: El archivo `.env` contiene credenciales sensibles. No debe ser commiteado al repositorio.

### 3.3 Configuración Java

Ajustar la ruta base en `ChatIARunner.java` para producción:

```java
private static final String PROD_BASE_DIR = "C:\\ERP\\Chat IA ERP";
```

## 4. Estructura de Archivos

```
Chat IA ERP/
├── java/
│   └── ChatIARunner.java          (Orquestador Java)
├── python/
│   ├── chat_ia_erp.py             (Motor de IA)
│   ├── db_schema.py               (Esquema de BD)
│   ├── requirements.txt           (Dependencias)
│   ├── .env                       (API Key - no commitear)
│   ├── instalar_entorno.ps1       (Script instalación)
│   └── .venv/                     (Entorno virtual)
└── docs/
    └── GUIA_CHAT_IA.md            (Esta guía)
```

## 5. Uso

### 5.1 Desde Java (Integración con SOLID)

```java
// Ejemplo de invocación desde SOLID
String consulta = "¿Cuántas legalizaciones hay este mes?";
ChatIARunner.main(new String[]{consulta});
```

### 5.2 Desde Línea de Comandos (Pruebas)

```bash
# Compilar Java
javac ChatIARunner.java

# Ejecutar con consulta
java ChatIARunner "¿Cuántas legalizaciones hay este mes?"
```

### 5.3 Directamente con Python (Debug)

```powershell
cd python
.venv\Scripts\python.exe chat_ia_erp.py "¿Cuántas legalizaciones hay este mes?"
```

## 6. Ejemplos de Consultas

### Consultas Simples
- "¿Cuántas legalizaciones hay este mes?"
- "Muestra las consignaciones del último mes"
- "¿Cuál es el total de gastos por empleado?"

### Consultas Complejas
- "Genera un reporte de gastos por empleado ordenado por monto descendente"
- "¿Cuál es la tendencia de legalizaciones en los últimos 3 meses?"
- "Muestra las legalizaciones con valor superior a 1 millón"

### Análisis
- "Analiza los gastos más altos por categoría"
- "¿Qué empleado tiene más consignaciones este año?"

## 7. Seguridad

El sistema implementa las siguientes medidas de seguridad:

1. **Validación de SQL**:
   - Solo permite consultas SELECT
   - Bloquea comandos peligrosos (DROP, DELETE, UPDATE, INSERT, etc.)
   - No permite acceso a esquemas del sistema

2. **Whitelist de Tablas**:
   - Solo permite consultar: legalizacion, linealegalizacion, consignacion

3. **Límites**:
   - Máximo 10,000 filas retornadas
   - Timeout de 30 segundos por query
   - Solo lectura (readonly session)

4. **API Key**:
   - Almacenada en archivo `.env` (no en código)
   - No debe ser commiteada al repositorio

## 8. Configuración de Base de Datos

El módulo se conecta a:
- **Host**: 192.168.0.21
- **Puerto**: 5432
- **Base de datos**: solid
- **Usuario**: postgres
- **Password**: AdminSolid2025

Para cambiar la configuración, editar la constante `DB_URI` en `chat_ia_erp.py`.

## 9. Troubleshooting

### Error: "GOOGLE_AI_API_KEY no encontrada"
- Verificar que el archivo `.env` existe en `python/`
- Verificar que contiene `GOOGLE_AI_API_KEY=tu_key_aqui`

### Error: "No se encuentra el entorno virtual"
- Ejecutar `instalar_entorno.ps1` en la carpeta `python/`
- Verificar que `.venv` fue creado correctamente

### Error: "Error en PostgreSQL"
- Verificar conectividad de red al servidor 192.168.0.21
- Verificar que PostgreSQL está corriendo
- Verificar credenciales de conexión

### Error: "Validación fallida"
- El SQL generado contiene comandos peligrosos
- Verificar que la consulta del usuario es apropiada
- Revisar logs para ver el SQL rechazado

## 10. Integración con SOLID

Para integrar con SOLID:

1. Compilar `ChatIARunner.java` y agregarlo al classpath de SOLID
2. Invocar desde SOLID pasando la consulta del usuario como argumento
3. Capturar la respuesta (stdout) y mostrarla al usuario

Ejemplo de integración:

```java
// En el código de SOLID
String consultaUsuario = obtenerConsultaDelUsuario();
String[] args = {consultaUsuario};
ChatIARunner.main(args);
// La respuesta se imprime en stdout
```

## 11. Modelo de IA

- **Modelo Principal**: Gemini 2.5 Flash
- **Modelos Fallback**: Gemini 2.0 Flash, Gemini 2.5 Flash Lite, Gemini 2.5 Pro
- **Temperatura**: 0.1 (generación SQL), 0.3 (análisis)
- **Max Tokens**: 1000 (SQL), 2000 (análisis)
- **Fallback Automático**: Si un modelo agota su cuota, automáticamente cambia al siguiente

Para cambiar los modelos, editar la lista `MODEL_PRIORITY` en `chat_ia_erp.py`.

## 12. Rendimiento

- Tiempo típico de respuesta: 4-10 segundos
- Depende de:
  - Complejidad de la consulta
  - Tamaño de resultados
  - Latencia de Google Gemini API
  - Latencia de red a PostgreSQL
  - Uso de modelos fallback (si aplica)

## 13. Mantenimiento

### Actualizar Dependencias
```powershell
cd python
.venv\Scripts\python.exe -m pip install --upgrade -r requirements.txt
```

### Verificar Estado
```powershell
# Verificar que Python funciona
.venv\Scripts\python.exe --version

# Verificar que Google Gemini SDK está instalado
.venv\Scripts\python.exe -c "import google.genai; print('Google Gemini OK')"
```

## 14. Soporte

Para problemas o preguntas, contactar al equipo de Mejoramiento Continuo.

