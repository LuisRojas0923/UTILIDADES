# Ejecuta este script desde la raíz del proyecto "INFORMES SOLID"
# para copiar todo el contenido importante a la carpeta Para_llevar.

$raiz = Split-Path -Parent $PSScriptRoot
$destino = $PSScriptRoot

Write-Host "Copiando desde: $raiz"
Write-Host "Hacia: $destino"

Copy-Item -Path "$raiz\Reporte estado cuenta" -Destination "$destino\Reporte estado cuenta" -Recurse -Force
Copy-Item -Path "$raiz\Reporte viaticos" -Destination "$destino\Reporte viaticos" -Recurse -Force
Copy-Item -Path "$raiz\.vscode" -Destination "$destino\.vscode" -Recurse -Force
Copy-Item -Path "$raiz\.classpath" -Destination "$destino\.classpath" -Force
Copy-Item -Path "$raiz\Reporte viaticos\.project" -Destination "$destino\.project" -Force

Write-Host "Listo. Revisa la carpeta Para_llevar."
