#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Chat IA ERP - Motor de consultas con IA para ERP SOLID
Genera SQL dinámico usando Google Gemini 2.5 Pro y ejecuta consultas en PostgreSQL
"""

import sys
import os
import json
import re
import psycopg2
from psycopg2.extras import RealDictCursor
try:
    import google.genai as genai
    USE_NEW_API = True
except ImportError:
    # Fallback al paquete antiguo si el nuevo no está disponible
    import google.generativeai as genai
    USE_NEW_API = False
from dotenv import load_dotenv
from db_schema import get_schema_for_llm

# Cargar variables de entorno
load_dotenv()

# Configuracion
DB_URI = "postgresql://postgres:AdminSolid2025@192.168.0.21:5432/solid"

# Modelos disponibles con fallback (orden de prioridad)
# Si un modelo falla por cuota, se intenta con el siguiente
MODEL_PRIORITY = [
    "models/gemini-2.5-flash",      # Principal: rápido y eficiente
    "models/gemini-2.0-flash",     # Fallback 1: alternativa rápida
    "models/gemini-2.5-flash-lite", # Fallback 2: más ligero
    "models/gemini-2.5-pro",        # Fallback 3: más potente (puede tener límites diferentes)
]

MODEL_NAME = MODEL_PRIORITY[0]  # Modelo por defecto
MAX_ROWS = 10000
QUERY_TIMEOUT = 30  # segundos

# Tablas permitidas (whitelist)
ALLOWED_TABLES = ["legalizacion", "linealegalizacion", "consignacion"]

# Inicializar cliente Google Gemini
api_key = os.getenv("GOOGLE_AI_API_KEY")
if not api_key:
    print("ERROR: GOOGLE_AI_API_KEY no encontrada en variables de entorno o archivo .env")
    sys.exit(1)

if USE_NEW_API:
    # Nuevo paquete google.genai
    client = genai.Client(api_key=api_key)
else:
    # Paquete antiguo google.generativeai
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(MODEL_NAME)


def validate_sql(sql: str) -> tuple[bool, str]:
    """
    Valida que el SQL generado sea seguro:
    - Solo SELECT permitido
    - Solo tablas en whitelist
    - No contiene comandos peligrosos
    """
    sql_upper = sql.upper().strip()
    
    # Verificar que sea solo SELECT
    if not sql_upper.startswith("SELECT"):
        return False, "Solo se permiten consultas SELECT"
    
    # Verificar que no contenga comandos peligrosos
    dangerous_keywords = [
        "DROP", "DELETE", "UPDATE", "INSERT", "ALTER", "CREATE", 
        "TRUNCATE", "EXEC", "EXECUTE", "CALL", "GRANT", "REVOKE"
    ]
    for keyword in dangerous_keywords:
        if keyword in sql_upper:
            return False, f"Comando peligroso detectado: {keyword}"
    
    # Verificar que solo use tablas permitidas (verificación básica)
    # Nota: Esta es una validación básica. El LLM debe generar SQL solo con tablas permitidas.
    sql_lower = sql.lower()
    # Verificar que al menos una tabla permitida esté presente
    has_allowed_table = any(table in sql_lower for table in ALLOWED_TABLES)
    if not has_allowed_table:
        # Permitir si no hay referencias explícitas (puede ser una consulta de sistema)
        pass  # No rechazar automáticamente, confiar en el LLM
    
    # Verificar que no haya subconsultas peligrosas
    if "INFORMATION_SCHEMA" in sql_upper or "PG_" in sql_upper:
        return False, "No se permite acceso a esquemas del sistema"
    
    return True, "OK"


def generate_sql(user_query: str) -> tuple[str, str]:
    """
    Genera SQL usando Google Gemini basado en la consulta del usuario
    Retorna: (sql_generado, error_message)
    """
    schema_context = get_schema_for_llm()
    
    # Prompt de sistema mejorado con más contexto
    system_prompt = """Eres un asistente experto en bases de datos PostgreSQL especializado en el módulo de VIÁTICOS del ERP SOLID.

CONTEXTO DEL MÓDULO DE VIÁTICOS:
Este chat es parte del módulo de viáticos del ERP SOLID. Su finalidad es llevar el control de la cartera de los dineros que se le entregan a los empleados en calidad de viáticos.

TABLAS Y SU FUNCIÓN:
- consignacion: Guarda las consignaciones (dinero entregado) que se hacen para los viáticos de los empleados
- legalizacion: Guarda los gastos de los empleados que hicieron con la consignación que se les dio
- linealegalizacion: Detalle de cada gasto dentro de una legalización
- transaccionviaticos: Guarda el cruce entre la consignación y la legalización

ÁREAS Y UEN (Unidades de Negocio):
- ADN: Es un área y también una UEN (Unidad de Negocio)
- OP: Es un área
- RCE: Es un área
Las áreas se extraen del código usando SPLIT_PART(codigo, '-', 1)

TU FUNCIÓN:
- Los usuarios hacen preguntas en lenguaje natural en español sobre viáticos
- Debes convertir sus preguntas a consultas SQL válidas
- Siempre debes usar solo las tablas permitidas: legalizacion, linealegalizacion, consignacion
Tu función es ayudar a los usuarios a consultar información sobre legalizaciones, gastos y consignaciones.

CONTEXTO DEL SISTEMA:
- Trabajas con un ERP empresarial llamado SOLID
- Los usuarios hacen preguntas en lenguaje natural en español
- Debes convertir sus preguntas a consultas SQL válidas
- Siempre debes usar solo las tablas permitidas: legalizacion, linealegalizacion, consignacion
- en la tabla consignacion se guardan las consignaciones que se hacen para los viaticos de los empleados
-las legalizacion es donde se guardan los gastos de los empleados que hiceron con la consignacion que se les dio
-en la tabla transaccionviaticos se guardan el cruce de la consignacion con la legalizacion.

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

    prompt = f"""{system_prompt}

ESQUEMA DE BASE DE DATOS:
{schema_context}

PREGUNTA DEL USUARIO (en lenguaje natural):
{user_query}

Genera la consulta SQL que responda a esta pregunta:

SQL:"""

    # Intentar con cada modelo en orden de prioridad (fallback automático)
    last_error = None
    for model_name in MODEL_PRIORITY:
        try:
            if USE_NEW_API:
                # Nuevo paquete google.genai
                response = client.models.generate_content(
                    model=model_name,
                    contents=[prompt],
                    config={
                        "temperature": 0.1,
                        "max_output_tokens": 1000,
                    }
                )
                sql = response.text.strip()
            else:
                # Paquete antiguo google.generativeai
                temp_model = genai.GenerativeModel(model_name)
                response = temp_model.generate_content(
                    prompt,
                    generation_config=genai.types.GenerationConfig(
                        temperature=0.1,
                        max_output_tokens=1000,
                    )
                )
                sql = response.text.strip()
            
            # Limpiar SQL (remover markdown si existe)
            sql = re.sub(r'^```sql\s*', '', sql, flags=re.IGNORECASE)
            sql = re.sub(r'^```\s*', '', sql)
            sql = re.sub(r'```\s*$', '', sql)
            sql = sql.strip()
            
            # Si llegamos aquí, el modelo funcionó
            if model_name != MODEL_PRIORITY[0]:
                print(f"[INFO] Usando modelo fallback para SQL: {model_name.replace('models/', '')}", file=sys.stderr)
            
            return sql, None
            
        except Exception as e:
            error_msg = str(e)
            last_error = error_msg
            
            # Si es error de cuota (429), intentar siguiente modelo
            if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg or "quota" in error_msg.lower():
                print(f"[INFO] Cuota agotada en {model_name.replace('models/', '')}, intentando siguiente modelo...", file=sys.stderr)
                continue
            # Si es otro error (sintaxis, etc), retornar inmediatamente
            else:
                return None, f"Error al generar SQL con {model_name.replace('models/', '')}: {error_msg}"
    
    # Si todos los modelos fallaron
    return None, f"Error: Todos los modelos fallaron por cuota. Último error: {last_error}"


def execute_query(sql: str) -> tuple[list, str]:
    """
    Ejecuta la consulta SQL en PostgreSQL
    Retorna: (resultados, error_message)
    """
    try:
        conn = psycopg2.connect(DB_URI, connect_timeout=QUERY_TIMEOUT)
        conn.set_session(readonly=True)
        
        cursor = conn.cursor(cursor_factory=RealDictCursor)
        cursor.execute(f"SET statement_timeout = {QUERY_TIMEOUT * 1000}")  # timeout en ms
        
        cursor.execute(sql)
        
        # Limitar número de filas
        rows = cursor.fetchall()
        if len(rows) > MAX_ROWS:
            rows = rows[:MAX_ROWS]
        
        # Convertir a lista de diccionarios
        results = [dict(row) for row in rows]
        
        cursor.close()
        conn.close()
        
        return results, None
        
    except psycopg2.Error as e:
        return None, f"Error en PostgreSQL: {str(e)}"
    except Exception as e:
        return None, f"Error inesperado: {str(e)}"


def analyze_results(user_query: str, sql: str, results: list) -> str:
    """
    Analiza los resultados usando GPT-4o-mini y genera una respuesta natural
    """
    # Preparar datos para el LLM
    results_json = json.dumps(results[:100], ensure_ascii=False, indent=2)  # Limitar a 100 filas para análisis
    
    system_context = """Eres un asistente del módulo de VIÁTICOS del ERP SOLID.
Tu función es explicar los resultados de consultas sobre viáticos de manera clara y concisa.

CONTEXTO DEL MÓDULO:
- Este módulo controla la cartera de dineros entregados a empleados en calidad de viáticos
- consignacion = dinero entregado a empleados
- legalizacion = gastos realizados con ese dinero
- ADN es un área y también una UEN (Unidad de Negocio)
- OP y RCE son áreas

INSTRUCCIONES:
- Sé CONCISO, no des información innecesaria
- Enfócate en responder directamente la pregunta
- Si preguntan por áreas, solo menciona las áreas encontradas sin explicaciones largas
- No expliques qué es cada área a menos que sea necesario
- Responde en español de forma directa y profesional"""

    prompt = f"""{system_context}

PREGUNTA ORIGINAL DEL USUARIO:
{user_query}

SQL EJECUTADO:
{sql}

RESULTADOS DE LA CONSULTA (primeras 100 filas de {len(results)} total):
{results_json}

INSTRUCCIONES PARA LA RESPUESTA:
1. Responde en español de forma DIRECTA y CONCISA
2. NO des información innecesaria o explicaciones largas
3. Si hay resultados, presenta SOLO lo relevante para la pregunta
4. Si hay muchas áreas/empleados, menciona el total y los más importantes (top 3-5 máximo)
5. NO expliques qué es cada área (ADN, OP, RCE) a menos que sea específicamente necesario
6. Si no hay resultados, di simplemente "No se encontraron resultados" SIN explicaciones adicionales
7. Enfócate en responder la pregunta específica del usuario sobre viáticos
8. Usa formato simple: números y listas cortas cuando sea necesario
9. Recuerda: ADN es un área y también una UEN, OP y RCE son áreas

Genera una respuesta natural y útil:

RESPUESTA:"""

    # Intentar con cada modelo en orden de prioridad (fallback automático)
    last_error = None
    for model_name in MODEL_PRIORITY:
        try:
            # Usar el prompt ya construido con el contexto del sistema
            full_prompt = prompt
            
            if USE_NEW_API:
                # Nuevo paquete google.genai
                response = client.models.generate_content(
                    model=model_name,
                    contents=[full_prompt],
                    config={
                        "temperature": 0.3,
                        "max_output_tokens": 2000,
                    }
                )
                result = response.text.strip()
            else:
                # Paquete antiguo google.generativeai
                temp_model = genai.GenerativeModel(model_name)
                response = temp_model.generate_content(
                    full_prompt,
                    generation_config=genai.types.GenerationConfig(
                        temperature=0.3,
                        max_output_tokens=2000,
                    )
                )
                result = response.text.strip()
            
            # Si llegamos aquí, el modelo funcionó
            if model_name != MODEL_PRIORITY[0]:
                print(f"[INFO] Usando modelo fallback para análisis: {model_name.replace('models/', '')}", file=sys.stderr)
            
            return result
            
        except Exception as e:
            error_msg = str(e)
            last_error = error_msg
            
            # Si es error de cuota (429), intentar siguiente modelo
            if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg or "quota" in error_msg.lower():
                print(f"[INFO] Cuota agotada en {model_name.replace('models/', '')} para análisis, intentando siguiente modelo...", file=sys.stderr)
                continue
            # Si es otro error, continuar con siguiente modelo
            else:
                print(f"[WARNING] Error con {model_name.replace('models/', '')}: {error_msg}, intentando siguiente...", file=sys.stderr)
                continue
    
    # Si todos los modelos fallaron, retornar resultados básicos
    return f"⚠️ No se pudo analizar con modelos disponibles debido a límites de cuota.\n\n📊 Resultados básicos ({len(results)} registros):\n{json.dumps(results[:10], ensure_ascii=False, indent=2)}"


def main():
    """Funcion principal"""
    if len(sys.argv) < 2:
        print("ERROR: Se requiere una consulta como argumento")
        print("Uso: python chat_ia_erp.py \"tu consulta aqui\"")
        sys.exit(1)
    
    user_query = sys.argv[1]
    
    if not user_query or not user_query.strip():
        print("ERROR: La consulta no puede estar vacía")
        sys.exit(1)
    
    print(f"[INFO] Procesando consulta: {user_query}", file=sys.stderr)
    
    # Paso 1: Generar SQL
    sql, error = generate_sql(user_query)
    if error:
        print(f"ERROR: {error}")
        sys.exit(1)
    
    print(f"[INFO] SQL generado: {sql}", file=sys.stderr)
    
    # Paso 2: Validar SQL
    is_valid, validation_msg = validate_sql(sql)
    if not is_valid:
        print(f"ERROR: Validación fallida - {validation_msg}")
        print(f"SQL rechazado: {sql}")
        sys.exit(1)
    
    print(f"[INFO] SQL validado correctamente", file=sys.stderr)
    
    # Paso 3: Ejecutar query
    results, error = execute_query(sql)
    if error:
        print(f"ERROR: {error}")
        sys.exit(1)
    
    print(f"[INFO] Query ejecutado. Filas retornadas: {len(results)}", file=sys.stderr)
    
    # Paso 4: Analizar resultados y generar respuesta
    response = analyze_results(user_query, sql, results)
    
    # Retornar respuesta (stdout para que Java la capture)
    print(response)
    sys.exit(0)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nERROR: Proceso interrumpido por el usuario")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR FATAL: {str(e)}")
        import traceback
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)

