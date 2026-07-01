@echo off
title Tunel publico Olnatura (puerto 8011)
cd /d "%~dp0"
echo.
echo === TUNEL OLNATURA (puerto 8011) ===
echo Requiere API corriendo: 1_API_OLNATURA.bat
echo.
where cloudflared >nul 2>&1
if errorlevel 1 (
    echo No hay cloudflared. Instale UNA VEZ:
    echo   winget install Cloudflare.cloudflared
    echo.
    echo O descargue: https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/
    pause
    exit /b 1
)
echo Copie la URL https que aparezca abajo (termina en trycloudflare.com)
echo Esa URL va en el codigo del Excel NIKZON.
echo Deje esta ventana abierta.
echo.
cloudflared tunnel --url http://127.0.0.1:8011
pause
