from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from contextlib import asynccontextmanager
import uvicorn
import time
import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime

# Forzar al script a usar la carpeta donde reside el .py como base
# Esto evita problemas cuando Windows lo ejecuta desde otra ubicacion
basedir = os.path.dirname(os.path.abspath(__file__))
os.chdir(basedir)
sys.path.insert(0, basedir)

# Forzar a Polars a saltar el chequeo de CPU para evitar crashes en servidores antiguos
os.environ["POLARS_SKIP_CPU_CHECK"] = "1"

# ==============================================================================
# CONFIGURACION DE LOGGING
# ==============================================================================
LOG_FILE = os.path.join(basedir, "sync.log")

# Crear logger
logger = logging.getLogger("SyncOT")
logger.setLevel(logging.INFO)

# Formato: [2026-01-08 14:30:15] INFO - Mensaje
formatter = logging.Formatter('[%(asctime)s] %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

# Handler para archivo (rotacion: max 5MB, guarda ultimos 3 archivos)
file_handler = RotatingFileHandler(LOG_FILE, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

# Handler para consola (para ver errores de inicio)
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(formatter)
console_handler.setLevel(logging.INFO)
logger.addHandler(console_handler)

# Importamos tu logica actual (pasamos el logger)
from upload_buffer_polars import upload_buffer_with_merge, set_logger

# Compartir el logger con el modulo de upload
set_logger(logger)

@asynccontextmanager
async def lifespan(app: FastAPI):
    """Maneja los eventos de inicio y cierre del servidor"""
    # Startup
    logger.info("=" * 60)
    logger.info("SERVIDOR FASTAPI INICIADO")
    logger.info(f"Puerto: 8099")
    logger.info(f"Log file: {LOG_FILE}")
    logger.info("=" * 60)
    yield
    # Shutdown (si es necesario agregar logica de limpieza)
    logger.info("SERVIDOR FASTAPI DETENIENDOSE")

app = FastAPI(
    title="Servicio de Sincronizacion OT - ERP SOLID",
    lifespan=lifespan
)

@app.post("/sync")
async def trigger_sync():
    """Endpoint para disparar la sincronizacion desde el ERP Java"""
    logger.info("=" * 60)
    logger.info("SOLICITUD DE SINCRONIZACION RECIBIDA")
    logger.info("=" * 60)
    
    try:
        start_time = time.time()
        # Ejecutamos la logica que ya tenemos validada
        upload_buffer_with_merge()
        elapsed = time.time() - start_time
        
        logger.info(f"SINCRONIZACION EXITOSA - Tiempo: {elapsed:.2f}s")
        logger.info("=" * 60)
        
        return {
            "status": "success",
            "message": "Sincronizacion completada exitosamente",
            "elapsed_seconds": round(elapsed, 2)
        }
    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        logger.error(f"ERROR EN SINCRONIZACION: {e}")
        logger.error(error_detail)
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
async def health_check():
    """Para verificar si el servidor esta vivo"""
    return {"status": "online", "server": "SOLID-ETL-SERVER"}

if __name__ == "__main__":
    # El servidor correra en el puerto 8099
    uvicorn.run(app, host="0.0.0.0", port=8099)

