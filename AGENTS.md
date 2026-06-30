# AI Agent Instructions for UTILIDADES

## Purpose
This repository is a collection of internal utilities for the Equipo de Mejoramiento Continuo. The AI agent should help with code navigation, bug fixes, improvements, and documentation, while respecting the repo's multi-language structure and production/experimental separation.

## Key Repo Areas
- `Chat IA ERP/`: Hybrid Java + Python module for intelligent ERP queries using Google Gemini. This is the main AI-related subsystem in this repo.
- `Python/`: Supporting scripts and references. The production ETL is not here; it is described in docs as belonging to `Servidor`.
- `VBA/`: Excel macros and database query macros for Tesorería, Viáticos, Inventario, and reports.
- `docs/`: Technical reports and analysis documents.
- Root-level utilities: Python scripts, PDF helpers, Excel analysis tools, and a small `package.json` for JS utilities.

## Important Conventions
- Prefer Python 3.10+ for Python work.
- The `Chat IA ERP` module uses `Chat IA ERP/python/.env` for sensitive API credentials. Do not add secrets into repo files.
- The `Chat IA ERP` system is designed for safe SELECT-only SQL generation on specific tables: `legalizacion`, `linealegalizacion`, `consignacion`.
- Do not assume `Python/` is the production ETL source. The repository README and Python README say the production ETL belongs to `Servidor` and `BufferUpload` is an outdated copy.
- When editing AI or SQL generation logic, respect exact column names and whitelist rules described in `Chat IA ERP/docs/DOCUMENTACION_TECNICA.md` and `Chat IA ERP/docs/GUIA_CHAT_IA.md`.

## How to Run
- `npm install` in the repo root if working on JavaScript utilities or `package-lock.json` dependencies.
- `cd "Chat IA ERP/python" && .\instalar_entorno.ps1` to set up the Python environment for the AI module.
- `cd "Chat IA ERP/java" && javac ChatIARunner.java` to compile the Java orchestrator.
- `java ChatIARunner "¿Cuántas legalizaciones hay este mes?"` to test the Chat IA ERP module from Java.
- For Python debugging in `Chat IA ERP/python`, use `.
    .venv\Scripts\python.exe chat_ia_erp.py "<pregunta>"`.

## Useful Documents
- `README.md` – repository overview and team context.
- `Python/README.md` – clarifies that production ETL is in `Servidor` and not the local Python copies.
- `Chat IA ERP/README.md` – quick start and module description.
- `Chat IA ERP/docs/GUIA_CHAT_IA.md` – deployment and usage guide.
- `Chat IA ERP/docs/DOCUMENTACION_TECNICA.md` – technical details, prompt rules, and database schema.
- `docs/INFORME_ANALISIS_JAVA_VS_PYTHON.md` – analysis of Java vs Python for ETL and environment notes.

## AI Behavior Notes
- When asked about repo structure, explain that this is a utilities repository with internal ERP tooling, not a monolithic application.
- Prioritize fixes in the active areas (`Chat IA ERP`, root scripts, `VBA/`) and avoid refactoring historical or legacy code unless the user requests it.
- If a task involves SQL or data access, ask the user to confirm which subsystem or dataset they want to target, since the repo contains multiple database-related modules.
- Leave deployment and environment-sensitive changes for explicit user approval.

## Recommended Next Customizations
- A dedicated skill for `Chat IA ERP` prompt and SQL generation rules.
- A hook for avoiding accidental `.env` or credentials check-ins.

> Note: This file is intended to help AI coding agents become productive quickly in this repo by summarizing structure, active areas, and important constraints.