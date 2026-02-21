# Script de instalacion automatica del entorno Chat IA ERP
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   INSTALADOR DE ENTORNO PYTHON PARA CHAT IA ERP" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Verificar si uv esta instalado
if (!(Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Host "[!] 'uv' no encontrado. Instalando uv..." -ForegroundColor Yellow
    powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
    $env:Path += ";$env:USERPROFILE\.cargo\bin"
}

# 2. Crear entorno virtual
Write-Host "[+] Creando entorno virtual en .venv..." -ForegroundColor Green
uv venv .venv

# 3. Instalar dependencias
Write-Host "[+] Instalando dependencias desde requirements.txt..." -ForegroundColor Green
uv pip install -r requirements.txt

Write-Host "`n[OK] Entorno listo para usar." -ForegroundColor White
Write-Host "Ruta del ejecutable: $(Get-Location)\.venv\Scripts\python.exe" -ForegroundColor Gray
Write-Host "`n[IMPORTANTE] No olvides crear el archivo .env con tu OPENAI_API_KEY" -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Cyan

