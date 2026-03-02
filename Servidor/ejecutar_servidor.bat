@echo off
cd /d "%~dp0"
set POLARS_SKIP_CPU_CHECK=1
".venv\Scripts\python.exe" api_server.py
pause
