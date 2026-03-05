# Documentación del Sistema de Reportes de Viáticos

Esta documentación explica el flujo de trabajo y la arquitectura del sistema implementado para la generación del **Reporte de Gastos de Viáticos**.

## Arquitectura del Sistema

El sistema se basa en una arquitectura de capas diseñada para ser escalable y profesional:

![Infografía del Flujo de Generación](C:/Users/amejoramiento6/.gemini/antigravity/brain/3167ee43-069c-4b10-9ab6-e6cbd55b2337/flujo_generacion_reporte_1768396873944.png)

### Capas de Software
- **Capa de Datos**: Gestionada por `ConexionDatos.java`, provee el acceso a PostgreSQL.
- **Capa Base (Skins)**: `BaseInforme.java` centraliza el diseño corporativo (Azul Navy, Gris F2F2F2).
- **Capa de Negocio**: `ReporteGastosViaticos.java` transforma filas de DB en celdas de PDF.

## Flujo de Generación del Reporte

A continuación se muestra el proceso desde que se solicita un reporte hasta que se abre el archivo PDF final:

```mermaid
graph TD
    A[Inicio: PruebaReporte.main] --> B{Conexión BD}
    B -- Exitosa --> C[Instanciar ReporteGastosViaticos]
    B -- Error --> Z[Fin con Error]
    
    C --> D[Llamar generar: CODIGO_RADICADO]
    
    subgraph "Proceso de Datos"
        D --> E[Consulta: Datos de Legalización]
        E --> F[Consulta: Primera OT del reporte]
        F --> G[Consulta: Cliente de la OT en otviaticos]
        G --> H[Sobrescribir campo 'OBRA' con nombre del Cliente]
    end
    
    H --> I[Inicializar Documento PDF: PageSize.LETTER]
    
    subgraph "Construcción del Reporte (iText)"
        I --> J[crearEncabezadoPrincipal: Logo + Título + Radicado]
        J --> K[crearTablaInfoEmpleado: Cabecera con Cliente/Obra]
        K --> L[crearBloqueCentral: Resumen OT y Rubros]
        L --> M[crearTablaDetalleGastos: Detalle + Cálculo de Totales]
    end
    
    M --> N[Cerrar Documento y Guardar Archivo PDF]
    N --> O[Llamar abrirArchivo: Abrir PDF automáticamente]
    O --> P[Fin del Proceso]
```

## Componentes Clave

| Archivo | Responsabilidad |
| :--- | :--- |
| `BaseInforme.java` | Define colores (`AZUL_NAVY`, `GRIS_CLARO`) y métodos de celdas. |
| `ReporteGastosViaticos.java` | Lógica principal, consultas SQL complejas y diseño del layout. |
| `PruebaReporte.java` | Punto de entrada para pruebas y desarrollo. |
| `itext-2.1.7.jar` | Librería core para la creación de archivos PDF. |

## Cómo Compilar y Ejecutar

Para compilar el sistema completo:
```powershell
javac -cp ".;itext-2.1.7.jar;postgresql-42.7.1.jar" *.java
```

Para generar un reporte de prueba:
```powershell
java -cp ".;itext-2.1.7.jar;postgresql-42.7.1.jar" PruebaReporte
```
