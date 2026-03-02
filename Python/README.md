# Python

La versión en producción del ETL (carga Excel → PostgreSQL) está en la carpeta **`/Servidor`** del repositorio.

- **Script principal:** `Servidor/upload_buffer_polars.py`
- **API (FastAPI):** `Servidor/api_server.py` — endpoint `POST /sync` para disparar la sincronización desde el ERP.

No uses la carpeta `BufferUpload` que estaba aquí; esa copia se eliminó para evitar duplicados. Todo el código activo está en `Servidor`.
