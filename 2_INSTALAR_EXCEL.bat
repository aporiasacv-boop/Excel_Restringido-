@echo off
title Instalar VBA en Excel NIKZON
cd /d "%~dp0"
echo.
echo === Instalar macros en LISTOS\ ===
echo Cierre Excel antes de continuar.
echo.
pause
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -Force
echo.
echo Listo. Archivos en carpeta LISTOS\
pause
