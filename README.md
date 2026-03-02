# UTILIDADES

Repositorio de utilidades del **Equipo de Mejoramiento Continuo** para integración con el ERP SOLID (PostgreSQL, Excel, VBA).

---

## Estructura del repositorio

| Carpeta | Descripción |
|---------|-------------|
| **Servidor** | Servicio de sincronización OT: API (FastAPI) que carga Excel desde carpetas de red a PostgreSQL (`basegeneralcostos` + `catalogoproducto`). Ver [Servidor/README.md](Servidor/README.md). |
| **Chat IA ERP** | Módulo de consultas en lenguaje natural al ERP (Java + Python + Gemini). |
| **Python** | Scripts y referencias; la versión en producción del ETL está en **Servidor**. |
| **VBA** | Macros Excel y consultas SQL (Tesorería, Viáticos, Establecimiento, Dataload, Informes). |
| **docs** | Informes técnicos (análisis Java vs Python para ETL, etc.). |

---

## Servicio de sincronización OT (Servidor)

- **API:** FastAPI en puerto **8099**.
- **Origen:** Excel en `\\192.168.0.3\...` (Base General, Postventa, MEMOFICHAS, Catálogo).
- **Destino:** PostgreSQL (tablas `basegeneralcostos`, `catalogoproducto`).
- **Despliegue:** Ver [Servidor/DESPLIEGUE.txt](Servidor/DESPLIEGUE.txt) y [Servidor/README.md](Servidor/README.md).

---

## Requisitos generales

- **Python:** 3.10+ (para Servidor y Chat IA ERP).
- **Git:** para clonar y actualizar el repositorio.

---

Uso interno - Equipo de Mejoramiento Continuo.
