#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script para listar modelos disponibles en Google Gemini API
"""

import google.genai as genai
import os
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("GOOGLE_AI_API_KEY")
if not api_key:
    print("ERROR: GOOGLE_AI_API_KEY no encontrada")
    exit(1)

client = genai.Client(api_key=api_key)

print("=" * 70)
print("MODELOS DISPONIBLES EN GOOGLE GEMINI API")
print("=" * 70)
print()

models = client.models.list()

# Separar modelos por tipo
embedding_models = []
gemini_models = []

for model in models:
    name = model.name
    if "embedding" in name.lower():
        embedding_models.append(name)
    elif "gemini" in name.lower():
        gemini_models.append(name)
    else:
        gemini_models.append(name)

print("MODELOS GEMINI (para generación de texto/SQL):")
print("-" * 70)
for i, model_name in enumerate(sorted(gemini_models), 1):
    # Extraer nombre corto
    short_name = model_name.replace("models/", "")
    print(f"  {i}. {short_name}")
    print(f"     Nombre completo: {model_name}")

print()
print("MODELOS DE EMBEDDING (para búsqueda semántica):")
print("-" * 70)
for i, model_name in enumerate(sorted(embedding_models), 1):
    short_name = model_name.replace("models/", "")
    print(f"  {i}. {short_name}")

print()
print("=" * 70)
print("RECOMENDACIONES:")
print("=" * 70)
print("  - Para consultas SQL rápidas: gemini-2.0-flash o gemini-2.5-flash")
print("  - Para análisis complejos: gemini-2.5-pro")
print("  - Para experimentación: gemini-2.0-flash-exp")
print("=" * 70)

