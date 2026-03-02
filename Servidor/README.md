# Servidor - Sincronización OT

API (FastAPI) que sincroniza datos desde archivos Excel en carpetas de red hacia PostgreSQL: tabla **basegeneralcostos** (Base General + Postventa + MEMOFICHAS) y tabla **catalogoproducto**.

---

## Inicio rápido

```cmd
cd Servidor
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python.exe api_server.py
```

- **Documentación interactiva:** http://localhost:8099/docs  
- **Health:** http://localhost:8099/health  
- **Disparar sincronización:** `POST http://localhost:8099/sync`

---

## Documentación incluida

| Archivo | Contenido |
|---------|-----------|
| [DESPLIEGUE.txt](DESPLIEGUE.txt) | Qué copiar al servidor, requisitos, instalación, arranque y detención. |
| [DOCKER_SMB.md](DOCKER_SMB.md) | Despliegue en Linux/Docker con montaje de carpetas SMB (192.168.0.3). |
| [MEJORAS_RENDIMIENTO.txt](MEJORAS_RENDIMIENTO.txt) | Análisis de rendimiento y posibles mejoras. |
| [create_table_basegeneralcostos.sql](create_table_basegeneralcostos.sql) | Esquema de referencia de la tabla. |

---

## Fuentes de datos (Excel)

| Fuente | Archivo / ruta | Tabla destino |
|--------|----------------|---------------|
| Base General | `\\192.168.0.3\Procesos Comunes SGI\Costos\...\BASE DE DATOS GENERAL.xlsm` | basegeneralcostos |
| Postventa | `\\192.168.0.3\Postventa\...\MTZ-SPT-02 Informe Gestion Postventa V1.xlsm` | basegeneralcostos |
| MEMOFICHAS | `\\192.168.0.3\Control Presupuestal\MEMOFICHA\CONSULTAS\CONSULTA MEMOFICHAS v2.xlsx` | basegeneralcostos |
| Catálogo | `\\192.168.0.3\Procesos Comunes SGI\Mejora\...\CATALOGO FINAL.xlsx` | catalogoproducto |

La columna **fuente** en `basegeneralcostos` indica el origen de cada fila: `BASE_GENERAL`, `POSTVENTA` o `MEMOFICHAS`. Solo se agregan órdenes que no existan ya en la base general.

---

## Scripts útiles

- **ejecutar_servidor.bat** — Arranca el API con el venv (ventana visible).
- **iniciar_servidor.pyw** — Arranca sin ventana (por ejemplo en el servidor).
- **detener_servidor.bat** — Finaliza el proceso que usa el puerto 8099.
- **instalar_entorno.ps1** — Crea `.venv` e instala dependencias (opcional, si se usa `uv`).

---

Equipo de Mejoramiento Continuo.
