@echo off
echo === Olnatura en PC de TI ===
echo.
echo 1. Deje esta ventana con la API corriendo
echo 2. En otra ventana abra los Excel desde LISTOS\
echo 3. Habilitar contenido y entrar con su usuario
echo.
cd /d "%~dp0api"
call INICIAR_API.bat
