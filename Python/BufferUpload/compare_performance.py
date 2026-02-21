"""
Script de Comparacion de Rendimiento: Python (Polars) vs Java (JDBC)

Ejecuta ambos loaders y compara los tiempos de ejecucion.
"""

import subprocess
import time
import os
import sys

# Rutas
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON_SCRIPT = os.path.join(SCRIPT_DIR, "upload_buffer_polars.py")
JAVA_PROJECT = os.path.join(SCRIPT_DIR, "..", "..", "Java", "ADBC_POC")

def run_python_loader():
    """Ejecuta el loader de Python y retorna el tiempo."""
    print("\n" + "=" * 70)
    print("  EJECUTANDO: Python (Polars)")
    print("=" * 70)
    
    start = time.time()
    
    result = subprocess.run(
        [sys.executable, PYTHON_SCRIPT],
        capture_output=False,
        cwd=SCRIPT_DIR
    )
    
    elapsed = time.time() - start
    return elapsed, result.returncode == 0


def run_java_loader():
    """Ejecuta el loader de Java y retorna el tiempo."""
    print("\n" + "=" * 70)
    print("  EJECUTANDO: Java (JDBC Batch)")
    print("=" * 70)
    
    start = time.time()
    
    # Compilar primero
    compile_result = subprocess.run(
        ["mvn", "clean", "package", "-q", "-DskipTests"],
        capture_output=True,
        cwd=JAVA_PROJECT,
        shell=True  # Necesario en Windows
    )
    
    if compile_result.returncode != 0:
        print("ERROR: Fallo la compilacion de Java")
        print(compile_result.stderr.decode())
        return 0, False
    
    # Ejecutar
    jar_path = os.path.join(JAVA_PROJECT, "target", "adbc-poc-1.0-SNAPSHOT.jar")
    result = subprocess.run(
        ["java", "-jar", jar_path],
        capture_output=False,
        cwd=JAVA_PROJECT
    )
    
    elapsed = time.time() - start
    return elapsed, result.returncode == 0


def main():
    print("\n" + "#" * 70)
    print("#  COMPARACION DE RENDIMIENTO: Python vs Java")
    print("#" * 70)
    
    results = {}
    
    # Ejecutar Python
    try:
        python_time, python_ok = run_python_loader()
        results["Python (Polars)"] = python_time if python_ok else None
    except Exception as e:
        print(f"Error ejecutando Python: {e}")
        results["Python (Polars)"] = None
    
    # Ejecutar Java
    try:
        java_time, java_ok = run_java_loader()
        results["Java (JDBC)"] = java_time if java_ok else None
    except Exception as e:
        print(f"Error ejecutando Java: {e}")
        results["Java (JDBC)"] = None
    
    # Mostrar resultados
    print("\n" + "=" * 70)
    print("  RESULTADOS DE COMPARACION")
    print("=" * 70)
    
    for name, elapsed in results.items():
        if elapsed is not None:
            print(f"  {name:20s}: {elapsed:8.2f} segundos")
        else:
            print(f"  {name:20s}: ERROR")
    
    # Determinar ganador
    valid_results = {k: v for k, v in results.items() if v is not None}
    if len(valid_results) == 2:
        winner = min(valid_results, key=valid_results.get)
        diff = abs(results["Python (Polars)"] - results["Java (JDBC)"])
        ratio = max(valid_results.values()) / min(valid_results.values())
        
        print(f"\n  GANADOR: {winner}")
        print(f"  Diferencia: {diff:.2f} segundos ({ratio:.1f}x mas rapido)")
    
    print("=" * 70)


if __name__ == "__main__":
    main()

