@echo off
title Tunel publico Olnatura (puerto 8011)
cd /d "%~dp0"
set "CF=%~dp0tools\cloudflared.exe"

echo.
echo === TUNEL OLNATURA (puerto 8011) ===
echo Requiere API corriendo: 1_API_OLNATURA.bat
echo.

if exist "%CF%" goto run
where cloudflared >nul 2>&1
if not errorlevel 1 (
    set "CF=cloudflared"
    goto run
)

echo No hay cloudflared.
echo Doble clic primero en: 0_INSTALAR_CLOUDFLARE.bat
pause
exit /b 1

:run
echo Copie la URL https que aparezca abajo (trycloudflare.com)
echo Esa URL va en el codigo del Excel NIKZON.
echo Deje esta ventana abierta.
echo.
"%CF%" tunnel --url http://127.0.0.1:8011
pause
