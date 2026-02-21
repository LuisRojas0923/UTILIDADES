# Documentación Técnica: Chat IA ERP

## 1. Resumen del Proyecto

Se ha implementado un módulo de chat con inteligencia artificial para el ERP SOLID que permite realizar consultas en lenguaje natural sobre las tablas de legalizaciones, consignaciones y gastos. El sistema utiliza Google Gemini 2.5 Flash para generar consultas SQL dinámicas y analizar resultados.

### Arquitectura

- **Java**: Orquestador e integración con SOLID
- **Python**: Motor de IA con Google Gemini API
- **PostgreSQL**: Base de datos con tablas legalizacion, linealegalizacion, consignacion

## 2. Estructura del Proyecto

```
Chat IA ERP/
├── java/
│   └── ChatIARunner.java          (Orquestador Java)
├── python/
│   ├── chat_ia_erp.py             (Motor de IA principal)
│   ├── db_schema.py               (Esquema de BD)
│   ├── requirements.txt           (Dependencias)
│   ├── instalar_entorno.ps1       (Script instalación)
│   ├── .env                       (API Key)
│   └── .venv/                     (Entorno virtual)
└── docs/
    ├── GUIA_CHAT_IA.md            (Guía de usuario)
    └── DOCUMENTACION_TECNICA.md   (Este documento)
```

## 3. Modelos de IA Utilizados

### Configuración de Modelos con Fallback Automático

El sistema implementa un mecanismo de fallback automático entre modelos para garantizar disponibilidad:

```python
MODEL_PRIORITY = [
    "models/gemini-2.5-flash",      # Principal: rápido y eficiente
    "models/gemini-2.0-flash",      # Fallback 1: alternativa rápida
    "models/gemini-2.5-flash-lite", # Fallback 2: más ligero
    "models/gemini-2.5-pro",        # Fallback 3: más potente
]
```

**Funcionamiento del Fallback:**
- Si un modelo agota su cuota (error 429), automáticamente intenta con el siguiente
- Transparente para el usuario (solo logs informativos)
- Garantiza continuidad del servicio

## 4. Prompts Utilizados

### 4.1 Prompt para Generación de SQL

```python
system_prompt = """Eres un asistente experto en bases de datos PostgreSQL especializado en el ERP SOLID.
Tu función es ayudar a los usuarios a consultar información sobre legalizaciones, gastos y consignaciones.

CONTEXTO DEL SISTEMA:
- Trabajas con un ERP empresarial llamado SOLID
- Los usuarios hacen preguntas en lenguaje natural en español
- Debes convertir sus preguntas a consultas SQL válidas
- Siempre debes usar solo las tablas permitidas: legalizacion, linealegalizacion, consignacion
- en la tabla consignacion se guardan las consignaciones que se hacen para los viaticos de los empleados
- las legalizacion es donde se guardan los gastos de los empleados que hiceron con la consignacion que se les dio
- en la tabla transaccionviaticos se guardan el cruce de la consignacion con la legalizacion.

REGLAS IMPORTANTES:
1. Si el usuario pregunta sobre "área", extrae el área usando SPLIT_PART:
   - Para consignaciones: SPLIT_PART(codigoconsignacion, '-', 1)
   - Para legalizaciones: SPLIT_PART(codigolegalizacion, '-', 1)
2. NOMBRES DE COLUMNAS EXACTOS (NO inventes nombres):
   - consignacion: codigoconsignacion (NO "codigo"), empleado, nombreempleado, valor, estado
   - legalizacion: codigo (PK), codigolegalizacion, empleado, nombreempleado, fechaaplicacion
   - linealegalizacion: legalizacion (FK), ot, centrocosto, subcentrocosto, categoria, valorconfactura, valorsinfactura, fecharealgasto
3. Para consignaciones, el código tiene formato "AREA-NUMERO" (ej: "ADN-C10020", "OP-C10008")
4. Para legalizaciones, el código tiene formato "AREA-NUMERO" en codigolegalizacion
5. Si preguntan por fechas, usa los campos fechaaplicacion, fecharealgasto según corresponda
6. Para valores monetarios, usa valorconfactura + valorsinfactura en legalizaciones
7. Para consignaciones, el valor está en el campo "valor"
8. Si preguntan por empleados, usa los campos empleado y nombreempleado
9. Siempre genera SQL válido y completo, sin dejar consultas incompletas
10. IMPORTANTE: En GROUP BY, usa la expresión completa, NO el alias. Ejemplo: GROUP BY SPLIT_PART(codigoconsignacion, '-', 1) NO GROUP BY area
11. Si usas funciones como SPLIT_PART, COUNT, SUM en SELECT, repite la misma expresión en GROUP BY

FORMATO DE RESPUESTA:
- Genera SOLO la consulta SQL
- Sin explicaciones adicionales
- Sin markdown (```sql)
- Sin comillas alrededor del SQL
- SQL completo y ejecutable"""
```

### 4.2 Prompt para Análisis de Resultados

```python
system_context = """Eres un asistente experto en análisis de datos empresariales del ERP SOLID.
Tu función es explicar los resultados de consultas de manera clara y útil para usuarios de negocio.

CONTEXTO:
- Trabajas con datos de legalizaciones, gastos y consignaciones
- Los usuarios necesitan respuestas claras y accionables
- Debes interpretar los datos en contexto empresarial
- Si hay áreas mencionadas (ADN, OP, RCE), explícalas en contexto de la empresa

INSTRUCCIONES PARA LA RESPUESTA:
1. Responde en español de forma natural, clara y profesional
2. Si hay resultados, presenta un resumen ejecutivo útil
3. Destaca los hallazgos más importantes primero
4. Si hay muchos resultados, menciona el total y presenta los top 5-10 más relevantes
5. Incluye números específicos y porcentajes cuando sea relevante
6. Si hay áreas (ADN, OP, RCE), explícalas en contexto empresarial
7. Si no hay resultados, explica por qué podría ser (filtros muy restrictivos, datos no disponibles, etc.)
8. Sé conciso pero completo - los usuarios necesitan información accionable
9. Usa formato claro: listas con viñetas para múltiples elementos, números destacados para totales"""
```

## 5. Esquema de Base de Datos

### 5.1 Tabla: legalizacion

```sql
Campos principales:
  - codigo (PK): Código único de la legalización
  - codigolegalizacion: Código de radicado (formato: AREA-NUMERO)
  - empleado: Documento de identidad del empleado
  - nombreempleado: Nombre del empleado
  - fechaaplicacion: Fecha de aplicación/entrega del reporte

Relaciones:
  - Uno a muchos con linealegalizacion (legalizacion.codigo = linealegalizacion.legalizacion)
```

### 5.2 Tabla: linealegalizacion

```sql
Campos principales:
  - legalizacion (FK): Referencia a legalizacion.codigo
  - ot: Número de orden de trabajo (puede estar vacío)
  - centrocosto: Centro de costo (usado si ot está vacío)
  - subcentrocosto: Subcentro de costo
  - categoria: Descripción/categoría del gasto
  - valorconfactura: Valor con factura
  - valorsinfactura: Valor sin factura
  - fecharealgasto: Fecha real del gasto

Reglas de negocio:
  - Si ot está vacío o NULL, se usa 'C' || centrocosto como identificador
  - El valor total aprobado es: valorconfactura + valorsinfactura
```

### 5.3 Tabla: consignacion

```sql
Campos principales:
  - codigoconsignacion: Código único de la consignación (contrato) - Formato: "AREA-NUMERO"
  - empleado: Documento de identidad del empleado
  - nombreempleado: Nombre del empleado
  - valor: Valor de la consignación
  - estado: Estado de la consignación (ej: 'CONTABILIZADO')

Reglas de negocio:
  - Impuesto 4x1000: valor * 0.004
  - Total consignación: valor + (valor * 0.004)
  - Para extraer área: SPLIT_PART(codigoconsignacion, '-', 1)
```

## 6. Resultados de Pruebas

### 6.1 Prueba 1: Consulta Simple - Conteo de Consignaciones

**Consulta del usuario:**
```
¿Cuántas consignaciones hay en total?
```

**SQL Generado:**
```sql
SELECT COUNT(*) FROM consignacion
```

**Resultado:**
- 78 consignaciones encontradas
- Tiempo de ejecución: ~7-8 segundos
- Estado: ✅ Éxito

**Respuesta del sistema:**
```
El número total de consignaciones registradas es de 78.
```

### 6.2 Prueba 2: Consulta Compleja - Agrupación por Empleado

**Consulta del usuario:**
```
Muestra un resumen de las consignaciones agrupadas por empleado con el total de cada uno
```

**SQL Generado:**
```sql
SELECT
  c.empleado,
  c.nombreempleado,
  SUM(c.valor) AS total_consignaciones
FROM consignacion c
GROUP BY
  c.empleado,
  c.nombreempleado
ORDER BY
  total_consignaciones DESC;
```

**Resultado:**
- 73 empleados encontrados
- Top 5 empleados con mayores consignaciones:
  - ECHAVARRIA LOAIZA JUAN DIEGO: $5,602,740.0
  - RICARDO RAMIREZ DAVID JOSE: $3,764,555.0
  - PACHECO JIMENEZ YERLIS ANTONIO: $3,320,650.0
  - COY PINZON ELIZABETH: $2,427,355.0
  - ZUÑIGA CARDONA JAMEZ: $2,000,000.0
- Tiempo de ejecución: ~7-8 segundos
- Estado: ✅ Éxito

### 6.3 Prueba 3: Consulta por Área

**Consulta del usuario:**
```
¿Qué área tiene más consignaciones?
```

**SQL Generado:**
```sql
SELECT SPLIT_PART(codigoconsignacion, '-', 1) AS area, COUNT(*) AS total 
FROM consignacion 
GROUP BY SPLIT_PART(codigoconsignacion, '-', 1) 
ORDER BY total DESC
```

**Resultado:**
- 3 áreas encontradas:
  - **ADN**: 64 consignaciones (mayor cantidad)
  - **OP**: 13 consignaciones
  - **RCE**: 1 consignación
- Tiempo de ejecución: ~4-6 segundos
- Estado: ✅ Éxito

**Respuesta del sistema:**
```
El área con más consignaciones es ADN con 64 consignaciones, seguida por OP con 13 y RCE con 1.
```

### 6.4 Prueba 4: Consulta sobre Áreas que han Solicitado Consignaciones

**Consulta del usuario:**
```
que areas han solicitado consignaciones?
```

**SQL Generado:**
```sql
SELECT SPLIT_PART(codigoconsignacion, '-', 1) AS area 
FROM consignacion 
GROUP BY SPLIT_PART(codigoconsignacion, '-', 1)
```

**Resultado:**
- 3 áreas identificadas: ADN, OP, RCE
- Análisis contextual incluido explicando cada área
- Tiempo de ejecución: ~4-6 segundos
- Estado: ✅ Éxito

## 7. Sistema de Seguridad

### 7.1 Validación de SQL

El sistema implementa validación estricta de SQL generado:

```python
def validate_sql(sql: str) -> tuple[bool, str]:
    """
    Valida que el SQL generado sea seguro:
    - Solo SELECT permitido
    - Solo tablas en whitelist
    - No contiene comandos peligrosos
    """
```

**Reglas de validación:**
- ✅ Solo permite consultas SELECT
- ✅ Bloquea comandos peligrosos: DROP, DELETE, UPDATE, INSERT, ALTER, CREATE, TRUNCATE, EXEC, etc.
- ✅ Whitelist de tablas: solo legalizacion, linealegalizacion, consignacion
- ✅ Bloquea acceso a esquemas del sistema (INFORMATION_SCHEMA, PG_*)

### 7.2 Límites de Seguridad

- **Timeout de queries**: 30 segundos máximo
- **Límite de filas**: 10,000 filas máximo
- **Sesión de solo lectura**: `conn.set_session(readonly=True)`

## 8. Configuración

### 8.1 Variables de Entorno

Archivo: `python/.env`
```env
GOOGLE_AI_API_KEY=AIzaSyAAUESKW2_HA5FVH5zq3X0Lg_VKEN56v6M
```

### 8.2 Conexión a Base de Datos

```python
DB_URI = "postgresql://postgres:AdminSolid2025@192.168.0.21:5432/solid"
```

### 8.3 Configuración de Modelos

```python
MODEL_NAME = "models/gemini-2.5-flash"  # Modelo por defecto
MODEL_PRIORITY = [
    "models/gemini-2.5-flash",
    "models/gemini-2.0-flash",
    "models/gemini-2.5-flash-lite",
    "models/gemini-2.5-pro",
]
```

## 9. Flujo de Ejecución

```
1. Usuario hace consulta (consola o argumento)
   ↓
2. Java (ChatIARunner.java) recibe consulta
   ↓
3. Java invoca Python mediante ProcessBuilder
   ↓
4. Python (chat_ia_erp.py):
   a. Genera SQL usando Gemini (con fallback automático)
   b. Valida SQL (seguridad)
   c. Ejecuta query en PostgreSQL
   d. Analiza resultados usando Gemini (con fallback automático)
   e. Genera respuesta natural en español
   ↓
5. Python retorna respuesta por stdout
   ↓
6. Java captura respuesta y la muestra al usuario
```

## 10. Características Implementadas

### 10.1 Fallback Automático entre Modelos

- ✅ Detección automática de errores de cuota (429)
- ✅ Cambio transparente entre modelos
- ✅ Logs informativos para debugging
- ✅ Resiliencia ante límites de API

### 10.2 Validación de SQL

- ✅ Solo SELECT permitido
- ✅ Whitelist de tablas
- ✅ Bloqueo de comandos peligrosos
- ✅ Validación de sintaxis básica

### 10.3 Análisis Inteligente

- ✅ Respuestas en español natural
- ✅ Contexto empresarial (áreas ADN, OP, RCE)
- ✅ Formato estructurado (listas, números destacados)
- ✅ Interpretación de resultados

### 10.4 Modo Interactivo

- ✅ Lectura desde consola si no hay argumentos
- ✅ Compatibilidad con argumentos de línea de comandos
- ✅ Detección automática de rutas

## 11. Rendimiento

### Tiempos Promedio

- **Consulta simple**: 4-6 segundos
- **Consulta compleja**: 7-10 segundos
- **Con fallback activado**: +2-3 segundos adicionales

### Factores que Afectan el Rendimiento

1. Complejidad de la consulta SQL generada
2. Tamaño de resultados retornados
3. Latencia de Google Gemini API
4. Latencia de red a PostgreSQL
5. Uso de modelos fallback (si aplica)

## 12. Limitaciones Conocidas

### 12.1 Cuotas de API

- Tier gratuito: límites por modelo (ej: 5-20 requests/minuto)
- El sistema maneja esto con fallback automático
- Para producción, considerar plan de pago

### 12.2 Validación de SQL

- Validación básica implementada
- No valida sintaxis completa de PostgreSQL
- Confía en que el LLM genere SQL válido

### 12.3 Contexto de Conversación

- No mantiene historial de conversación
- Cada consulta es independiente
- No recuerda consultas anteriores

## 13. Mejoras Futuras Sugeridas

1. **Historial de conversación**: Mantener contexto entre consultas
2. **Caché de consultas**: Evitar regenerar SQL para consultas similares
3. **Validación SQL más robusta**: Parser SQL completo
4. **Métricas y logging**: Tracking de uso y rendimiento
5. **Interfaz web**: UI amigable en lugar de solo consola
6. **Exportación de resultados**: Generar reportes en Excel/PDF

## 14. Dependencias

### Python

```
google-generativeai>=0.3.0
google-genai>=1.56.0
psycopg2-binary>=2.9.11
python-dotenv>=1.0.0
```

### Java

- JDK 8 o superior
- Sin dependencias externas (solo librerías estándar)

## 15. Ejemplos de Uso

### Desde Java (con argumentos)

```bash
java -cp "Chat IA ERP/java" ChatIARunner "¿Cuántas consignaciones hay?"
```

### Desde Java (modo interactivo)

```bash
java -cp "Chat IA ERP/java" ChatIARunner
# Ingrese su consulta: ¿Qué área tiene más consignaciones?
```

### Desde Python directamente (debug)

```bash
cd "Chat IA ERP/python"
.venv/Scripts/python.exe chat_ia_erp.py "¿Cuántas consignaciones hay?"
```

## 16. Troubleshooting

### Error: "Cuota agotada"
- **Solución**: El sistema cambia automáticamente a otro modelo
- **Si todos fallan**: Esperar unos minutos y reintentar

### Error: "No se encuentra el entorno virtual"
- **Solución**: Ejecutar `python/instalar_entorno.ps1`

### Error: "GOOGLE_AI_API_KEY no encontrada"
- **Solución**: Verificar que existe `python/.env` con la API key

### Error: "Error en PostgreSQL"
- **Solución**: Verificar conectividad de red y que PostgreSQL esté corriendo

## 17. Conclusión

El módulo Chat IA ERP está completamente funcional y listo para integración con SOLID. Implementa:

- ✅ Generación automática de SQL desde lenguaje natural
- ✅ Análisis inteligente de resultados
- ✅ Sistema de fallback automático entre modelos
- ✅ Validación de seguridad
- ✅ Modo interactivo y por argumentos
- ✅ Documentación completa

**Estado del proyecto**: ✅ Completado y probado

**Fecha de documentación**: Enero 2025

