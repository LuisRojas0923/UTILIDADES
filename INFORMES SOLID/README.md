# INFORMES SOLID

Módulo Java para generación de informes del ERP SOLID: reportes de gastos de viáticos (PDF), estado de cuenta (PDF y Excel vía API) y utilidades de conexión a PostgreSQL.

## Requisitos

- **Java 8** (JDK 1.8)
- Acceso de red a:
  - PostgreSQL del ERP (`192.168.0.21:5432`, base `solid`)
  - API de estado de cuenta en Excel (`EstadoCuentaConsumer`, solo si se usa ese flujo)
- Dependencias en `lib/`:
  - `itext-2.1.7.jar` — generación de PDF
  - `postgresql-42.7.1.jar` — driver JDBC

## Estructura del proyecto

```
INFORMES SOLID/
├── src/                        # Código activo principal (proyecto Eclipse)
│   ├── PruebaReporte.java      # Entrada: reporte de viáticos (PDF)
│   ├── ReporteGastosViaticos.java
│   ├── EstadoCuentaConsumer.java  # Entrada: estado de cuenta (Excel vía HTTP)
│   ├── BaseInforme.java
│   ├── ConexionDatos.java
│   ├── ImageConverter.java
│   └── Imagen.java
│
├── Reporte viaticos/           # Módulo completo de viáticos (versión extendida)
│   ├── RunnerReporte.java      # Entrada alternativa
│   ├── ReporteGastosViaticos.java
│   ├── BaseInforme.java
│   ├── ConexionDatos.java
│   └── ... (GeneraInforme, SOLID, MTabla, etc.)
│
├── Reporte estado cuenta/      # Módulo de estado de cuenta en PDF
│   ├── RunnerEstadoCuenta.java # Entrada
│   ├── ReporteEstadoCuenta.java
│   └── EstadoDeCuentaV2.sql    # Consulta SQL de referencia
│
├── Para_llevar/                # Copias portables para desplegar en otro entorno
│   ├── Reporte viaticos/
│   └── Reporte estado cuenta/
│
├── lib/                        # JARs (iText, PostgreSQL)
├── docs/                       # Documentación técnica adicional
│   └── DOCUMENTACION_FLUJO.md  # Flujo detallado del reporte de viáticos
├── images/                     # Logos y recursos gráficos
├── temp/                       # Salidas temporales (PDF, XLSX generados)
├── out/                        # Clases compiladas (salida manual)
└── bin/                        # Clases compiladas (salida Eclipse)
```

## Informes disponibles

| Informe | Formato | Carpeta / entrada | Descripción |
| :--- | :--- | :--- | :--- |
| Gastos de viáticos | PDF | `src/PruebaReporte.java` o `Reporte viaticos/RunnerReporte.java` | Consulta legalizaciones en PostgreSQL y genera PDF con iText |
| Estado de cuenta | PDF | `Reporte estado cuenta/RunnerEstadoCuenta.java` | Genera estado de cuenta por cédula |
| Estado de cuenta | Excel | `src/EstadoCuentaConsumer.java` | Descarga XLSX desde endpoint HTTP del ERP |

## Arquitectura (viáticos)

El reporte de viáticos sigue una arquitectura en capas:

- **Capa de datos** — `ConexionDatos.java`: conexión JDBC a PostgreSQL
- **Capa base (diseño)** — `BaseInforme.java`: colores corporativos y utilidades de celdas PDF
- **Capa de negocio** — `ReporteGastosViaticos.java`: consultas SQL, layout y generación del documento

Para el diagrama de flujo completo y la tabla de componentes, ver [`docs/DOCUMENTACION_FLUJO.md`](docs/DOCUMENTACION_FLUJO.md).

## Compilar y ejecutar

Desde la raíz de `INFORMES SOLID`:

### Reporte de viáticos (`src/`)

```powershell
cd "INFORMES SOLID"

# Compilar
javac -encoding UTF-8 -cp "lib/*" -d out src\*.java

# Ejecutar (editar el código de legalización en PruebaReporte.java antes)
java -cp "out;lib/*" PruebaReporte
```

### Estado de cuenta en Excel (API)

```powershell
javac -encoding UTF-8 -cp "lib/*" -d out src\EstadoCuentaConsumer.java
java -cp "out;lib/*" EstadoCuentaConsumer
```

El archivo se guarda en `temp/Reporte_<cedula>.xlsx`.

### Estado de cuenta en PDF

```powershell
cd "Reporte estado cuenta"
javac -encoding UTF-8 -cp "..\lib\*" *.java
java -cp ".;..\lib\*" RunnerEstadoCuenta
```

### Viáticos (módulo extendido en `Reporte viaticos/`)

```powershell
cd "Reporte viaticos"
javac -encoding UTF-8 -cp "..\lib\*" *.java
java -cp ".;..\lib\*" RunnerReporte
```

## Proyecto Eclipse

El directorio incluye `.project` y `.classpath`. La fuente oficial del IDE apunta a `src/` y la salida de compilación a `bin/`. Si se abre en Eclipse, importar como *Java Project* existente.

> **Nota:** `.classpath` referencia los JAR en la raíz del proyecto; los archivos actuales están en `lib/`. Ajustar el classpath en Eclipse o usar los comandos `javac` de arriba con `lib/*`.

## Carpeta `Para_llevar/`

Contiene copias autocontenidas de los módulos de **Reporte viaticos** y **Reporte estado cuenta** para copiar a otro equipo o entorno sin arrastrar todo el repositorio. No es la fuente principal de desarrollo; usar `src/` y las carpetas de módulo para cambios.

## Salidas generadas

- PDF de viáticos: se crean en la carpeta de trabajo y se abren automáticamente al finalizar (`ReporteGastosViaticos`)
- Excel de estado de cuenta: `temp/Reporte_<cedula>.xlsx`
- Ejemplos previos en la raíz y en `temp/` (p. ej. `Reporte_Viaticos_RCE-L1190.pdf`)

## Estado de la organización

El módulo funciona, pero hay **código duplicado** entre `src/`, `Reporte viaticos/` y `Para_llevar/`, además de dos carpetas de compilados (`out/` y `bin/`). Para mantenimiento futuro se recomienda:

1. Tratar `src/` como única fuente de verdad para el flujo activo
2. Mantener `Reporte estado cuenta/` como módulo separado hasta unificarlo
3. Usar `Para_llevar/` solo como empaquetado, no para editar
4. No versionar credenciales en código (hoy están en `ConexionDatos.java` y runners); mover a variables de entorno o archivo local ignorado por git

## Documentación relacionada

- [`docs/DOCUMENTACION_FLUJO.md`](docs/DOCUMENTACION_FLUJO.md) — flujo de generación del reporte de viáticos
- [`docs/compile_errors_*.txt`](docs/) — registros de errores de compilación históricos
- Repositorio padre: [`../AGENTS.md`](../AGENTS.md) — contexto general del equipo de Mejoramiento Continuo
