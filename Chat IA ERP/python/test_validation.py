#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de validacion basica para el modulo Chat IA ERP
Valida que los componentes principales esten correctamente configurados
"""

import os
import sys

def test_imports():
    """Valida que las dependencias esten instaladas"""
    print("[TEST] Validando imports...")
    try:
        import psycopg2
        print("  [OK] psycopg2 importado correctamente")
    except ImportError:
        print("  [ERROR] psycopg2 no esta instalado")
        return False
    
    try:
        from dotenv import load_dotenv
        print("  [OK] python-dotenv importado correctamente")
    except ImportError:
        print("  [ERROR] python-dotenv no esta instalado")
        return False
    
    try:
        from openai import OpenAI
        print("  [OK] openai importado correctamente")
    except ImportError:
        print("  [ERROR] openai no esta instalado")
        return False
    
    return True

def test_db_schema():
    """Valida que el esquema de BD se puede cargar"""
    print("[TEST] Validando db_schema.py...")
    try:
        from db_schema import get_schema_for_llm
        schema = get_schema_for_llm()
        if len(schema) > 100:
            print("  [OK] Esquema de BD cargado correctamente")
            return True
        else:
            print("  [ERROR] Esquema de BD parece estar vacio")
            return False
    except Exception as e:
        print(f"  [ERROR] No se pudo cargar db_schema: {e}")
        return False

def test_env_file():
    """Valida que existe archivo .env con API key"""
    print("[TEST] Validando archivo .env...")
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    if os.path.exists(env_path):
        print("  [OK] Archivo .env encontrado")
        # Verificar que tiene contenido
        with open(env_path, 'r') as f:
            content = f.read()
            if "OPENAI_API_KEY" in content and len(content) > 20:
                print("  [OK] OPENAI_API_KEY configurada")
                return True
            else:
                print("  [WARNING] OPENAI_API_KEY parece estar vacia o mal formateada")
                return False
    else:
        print("  [WARNING] Archivo .env no encontrado (crear desde .env.example)")
        return False

def test_sql_validation():
    """Valida la funcion de validacion SQL"""
    print("[TEST] Validando funcion de validacion SQL...")
    try:
        # Importar la funcion (simular import)
        sys.path.insert(0, os.path.dirname(__file__))
        from chat_ia_erp import validate_sql
        
        # Test 1: SELECT valido
        sql1 = "SELECT * FROM legalizacion"
        valid, msg = validate_sql(sql1)
        if valid:
            print("  [OK] SELECT valido aceptado")
        else:
            print(f"  [ERROR] SELECT valido rechazado: {msg}")
            return False
        
        # Test 2: DELETE peligroso
        sql2 = "DELETE FROM legalizacion"
        valid, msg = validate_sql(sql2)
        if not valid:
            print("  [OK] DELETE peligroso rechazado correctamente")
        else:
            print("  [ERROR] DELETE peligroso fue aceptado")
            return False
        
        # Test 3: DROP peligroso
        sql3 = "DROP TABLE legalizacion"
        valid, msg = validate_sql(sql3)
        if not valid:
            print("  [OK] DROP peligroso rechazado correctamente")
        else:
            print("  [ERROR] DROP peligroso fue aceptado")
            return False
        
        return True
        
    except Exception as e:
        print(f"  [ERROR] Error en validacion SQL: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Ejecuta todos los tests"""
    print("=" * 60)
    print("VALIDACION DEL MODULO CHAT IA ERP")
    print("=" * 60)
    print()
    
    tests = [
        ("Imports", test_imports),
        ("Esquema BD", test_db_schema),
        ("Archivo .env", test_env_file),
        ("Validacion SQL", test_sql_validation),
    ]
    
    results = []
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
        except Exception as e:
            print(f"  [ERROR] Excepcion en test {name}: {e}")
            results.append((name, False))
        print()
    
    # Resumen
    print("=" * 60)
    print("RESUMEN DE VALIDACION")
    print("=" * 60)
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for name, result in results:
        status = "[OK]" if result else "[FALLO]"
        print(f"{status} {name}")
    
    print()
    print(f"Total: {passed}/{total} tests pasados")
    
    if passed == total:
        print("\n[EXITO] Todos los tests pasaron!")
        return 0
    else:
        print("\n[ADVERTENCIA] Algunos tests fallaron. Revisar arriba.")
        return 1

if __name__ == "__main__":
    sys.exit(main())

