@echo off
title Instalar Cloudflare cloudflared (solo UNA vez)
cd /d "%~dp0"
set "CF=%~dp0tools\cloudflared.exe"

echo.
echo === 0 - INSTALAR CLOUDFLARE ===
echo.

if not exist "%~dp0tools" mkdir "%~dp0tools"

if exist "%CF%" (
    echo Ya esta instalado en:
    echo   %CF%
    echo.
    "%CF%" --version
    echo.
    echo Siguiente: 1_API_OLNATURA.bat y luego 2_TUNEL_OLNATURA.bat
    pause
    exit /b 0
)

echo Descargando cloudflared (gratis, sin cuenta)...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { $u='https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe'; Invoke-WebRequest -Uri $u -OutFile '%CF%' -UseBasicParsing; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"

if errorlevel 1 (
    echo.
    echo ERROR al descargar. Compruebe internet o ejecute como administrador.
    pause
    exit /b 1
)

if not exist "%CF%" (
    echo ERROR: no se creo el archivo.
    pause
    exit /b 1
)

echo.
echo Listo: %CF%
"%CF%" --version
echo.
echo Siguiente paso:
echo   1. Doble clic en 1_API_OLNATURA.bat
echo   2. Doble clic en 2_TUNEL_OLNATURA.bat
echo.
pause
