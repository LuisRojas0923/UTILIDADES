@echo off
echo Deteniendo servicio de sincronizacion...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8099') do (
    taskkill /PID %%a /F 2>nul
)
echo Servicio detenido.
pause
