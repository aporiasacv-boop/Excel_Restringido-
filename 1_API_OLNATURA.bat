@echo off
title Olnatura API - puerto 8011
cd /d "%~dp0api"
echo.
echo === 2/3 API OLNATURA (Python :8011) ===
echo Deje esta ventana abierta.
echo.

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py -3.12"
if not defined PY where py >nul 2>&1 && set "PY=py -3"

if not defined PY (
    echo No se encontro Python. Instale Python 3.12.
    pause
    exit /b 1
)

echo Usando: %PY%
%PY% -m pip install -r requirements.txt -q
echo Iniciando en http://127.0.0.1:8011 ...
%PY% -m uvicorn main:app --host 127.0.0.1 --port 8011
pause
