# ============================================================================
# LANZADOR SIN VENTANA - Servicio de Sincronizacion OT
# ============================================================================
import subprocess
import os

basedir = os.path.dirname(os.path.abspath(__file__))
api_script = os.path.join(basedir, "api_server.py")

# Preferir Python del venv de esta carpeta; si no existe, usar el del sistema
basedir_venv = os.path.join(basedir, ".venv", "Scripts", "python.exe")
if os.path.isfile(basedir_venv):
    python_exe = basedir_venv
else:
    python_exe = r"C:\Program Files\Python312\python.exe"

# Flag para ocultar ventana en Windows
CREATE_NO_WINDOW = 0x08000000

try:
    subprocess.Popen(
        [python_exe, api_script],
        cwd=basedir,
        creationflags=CREATE_NO_WINDOW,
        env={**os.environ, "POLARS_SKIP_CPU_CHECK": "1"},
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
except Exception as e:
    with open(os.path.join(basedir, "error_inicio.log"), "w") as f:
        f.write(f"Error al iniciar: {e}")
