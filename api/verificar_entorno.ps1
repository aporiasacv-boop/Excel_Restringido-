# Verifica Python, puertos y carpetas antes de instalar Olnatura Auth en PC de TI.
# No modifica nada. Solo lectura.

param(
    [int]$PuertoDeseado = 8010
)

Write-Host "=== Olnatura Auth - Verificacion de entorno ===" -ForegroundColor Cyan
Write-Host ""

# Python
Write-Host "--- Python ---" -ForegroundColor Yellow
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    Write-Host "  [FALTA] Python no encontrado en PATH" -ForegroundColor Red
    Write-Host "  Instalar: https://www.python.org/downloads/ (marcar 'Add to PATH')"
} else {
    $ver = python --version 2>&1
    Write-Host "  [OK] $ver en $($pythonCmd.Source)" -ForegroundColor Green
    Write-Host "  pip:" (python -m pip --version 2>&1)
}

Write-Host ""
Write-Host "--- Puertos en uso (TCP escuchando) ---" -ForegroundColor Yellow
$listening = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue |
    Select-Object LocalAddress, LocalPort, OwningProcess |
    Sort-Object LocalPort -Unique

if ($listening) {
    foreach ($conn in $listening) {
        $procName = (Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue).ProcessName
        Write-Host ("  Puerto {0,-6} -> {1} (PID {2})" -f $conn.LocalPort, $procName, $conn.OwningProcess)
    }
} else {
    Write-Host "  (no se pudo listar; pruebe: netstat -ano | findstr LISTENING)"
}

Write-Host ""
Write-Host "--- Puerto recomendado para Olnatura: $PuertoDeseado ---" -ForegroundColor Yellow
$ocupado = Get-NetTCPConnection -LocalPort $PuertoDeseado -State Listen -ErrorAction SilentlyContinue
if ($ocupado) {
    Write-Host "  [OCUPADO] Elegir otro puerto en api\.env (API_PORT=)" -ForegroundColor Red
} else {
    Write-Host "  [LIBRE] Puede usar API_PORT=$PuertoDeseado en api\.env" -ForegroundColor Green
}

Write-Host ""
Write-Host "--- Puertos tipicos de otros proyectos ---" -ForegroundColor Yellow
@(8000, 8001, 8080, 5000, 3000, 5432, 8010, 8011) | ForEach-Object {
    $p = $_
    $hit = Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue
    if ($hit) {
        $procName = (Get-Process -Id $hit[0].OwningProcess -ErrorAction SilentlyContinue).ProcessName
        Write-Host "  Puerto $p OCUPADO por $procName" -ForegroundColor DarkYellow
    }
}

Write-Host ""
Write-Host "--- Procesos uvicorn / cloudflared ---" -ForegroundColor Yellow
Get-Process -Name uvicorn, python, cloudflared -ErrorAction SilentlyContinue |
    ForEach-Object { Write-Host "  $($_.ProcessName) PID $($_.Id)" }

Write-Host ""
Write-Host "--- Carpeta de este proyecto ---" -ForegroundColor Yellow
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "  $here"
if (Test-Path (Join-Path $here "data\olnatura.db")) {
    Write-Host "  [OK] Base de datos local: data\olnatura.db" -ForegroundColor Green
} else {
    Write-Host "  [INFO] Aun no existe data\olnatura.db (se crea al iniciar la API)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "--- Dependencias Python (api) ---" -ForegroundColor Yellow
if ($pythonCmd) {
    Push-Location $here
    if (Test-Path "requirements.txt") {
        $needInstall = $false
        foreach ($line in Get-Content "requirements.txt") {
            if ($line -match "^([a-zA-Z0-9_-]+)") {
                $pkg = $Matches[1]
                if ($pkg -eq "uvicorn") { continue }
                pip show $pkg 2>$null | Out-Null
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "  [FALTA] $pkg" -ForegroundColor Red
                    $needInstall = $true
                }
            }
        }
        if (-not $needInstall) {
            Write-Host "  [OK] Dependencias de requirements.txt instaladas" -ForegroundColor Green
        } else {
            Write-Host "  Ejecutar: python -m pip install -r requirements.txt"
        }
    }
    Pop-Location
}

Write-Host ""
Write-Host "=== Fin. Olnatura usa carpeta propia (api/data/) y puerto configurable en .env ===" -ForegroundColor Cyan
