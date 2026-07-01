@echo off
title Instalar VBA en Excel NIKZON
cd /d "%~dp0"
echo.
echo === Instalar macros en LISTOS\ ===
echo.

if not exist "%~dp0entrada" mkdir "%~dp0entrada"

dir /b "%~dp0entrada\NIKZON*.xlsm" >nul 2>&1
if errorlevel 1 (
    dir /b "%~dp0NIKZON*.xlsm" >nul 2>&1
    if errorlevel 1 (
        echo ERROR: No hay archivos NIKZON*.xlsm
        echo.
        echo Copie los 4 Excel originales a la carpeta:
        echo   %~dp0entrada
        echo.
        echo Ejemplo: NIKZON 1.xlsm, NIKZON 2.xlsm ...
        echo Luego ejecute este .bat otra vez.
        pause
        exit /b 1
    )
)

tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if not errorlevel 1 (
    echo Cierre Excel y pulse una tecla para continuar...
    pause
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -Force
if errorlevel 1 (
    echo.
    echo La instalacion fallo. Revise el mensaje de arriba.
    pause
    exit /b 1
)

echo.
dir /b "%~dp0LISTOS\*.xlsm" 2>nul
if errorlevel 1 (
    echo AVISO: LISTOS\ sigue vacio. Revise permisos VBA en Excel.
) else (
    echo Archivos generados en LISTOS\
)
echo.
pause
