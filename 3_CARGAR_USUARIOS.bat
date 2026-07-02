@echo off
title Cargar usuarios desde usuarios\USUARIOS.txt
cd /d "%~dp0api"
echo.
echo === Cargar / actualizar usuarios ===
echo Archivo: %~dp0usuarios\USUARIOS.txt
echo.

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py -3.12"
if not defined PY where py >nul 2>&1 && set "PY=py -3"

if not defined PY (
    echo No se encontro Python.
    pause
    exit /b 1
)

%PY% cargar_usuarios.py
echo.
pause
