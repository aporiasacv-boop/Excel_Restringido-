# Lee la URL publica del tunel "olnatura" desde la API local de ngrok (puerto 4040)
# y actualiza src\modAppConstants.bas. Ejecutar con ngrok ya corriendo.

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$constants = Join-Path $root "src\modAppConstants.bas"

try {
    $resp = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 5
} catch {
    Write-Host "No se pudo leer ngrok en http://127.0.0.1:4040" -ForegroundColor Red
    Write-Host "Arranque primero INICIAR_NGROK.bat (en prueba_ov) o: ngrok start dynamics olnatura"
    exit 1
}

$tunnel = $resp.tunnels | Where-Object { $_.name -eq "olnatura" -or $_.name -eq "command_line" } | Select-Object -First 1
if (-not $tunnel) {
    $tunnel = $resp.tunnels | Where-Object { $_.config.addr -match ":8011" } | Select-Object -First 1
}
if (-not $tunnel -or -not $tunnel.public_url) {
    Write-Host "No hay tunel olnatura activo (puerto 8011)." -ForegroundColor Red
    $resp.tunnels | ForEach-Object { Write-Host "  - $($_.name): $($_.public_url) -> $($_.config.addr)" }
    exit 1
}

$url = ($tunnel.public_url -replace "/$", "").Trim()
if ($url -notmatch "^https://") {
    Write-Host "URL inesperada: $url" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $constants)) {
    Write-Host "No existe $constants" -ForegroundColor Red
    exit 1
}

$text = Get-Content $constants -Raw -Encoding Default
$newText = $text -replace '(Public Const API_BASE_URL As String = ")[^"]*(")', "`${1}$url`${2}"
if ($text -eq $newText) {
    Write-Host "API_BASE_URL ya era: $url" -ForegroundColor Yellow
} else {
    Set-Content $constants -Value $newText.TrimEnd() -Encoding Default
    Write-Host "Actualizado API_BASE_URL -> $url" -ForegroundColor Green
    Write-Host "Siguiente paso: cerrar Excel y ejecutar INSTALAR.bat en la raiz del proyecto."
}
