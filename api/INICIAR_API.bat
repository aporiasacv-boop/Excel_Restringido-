@echo off
cd /d "%~dp0"
if not exist .env copy .env.example .env

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py -3.12"
if not defined PY where py >nul 2>&1 && set "PY=py -3"
if not defined PY if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PY=%LocalAppData%\Programs\Python\Python312\python.exe"
if not defined PY if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PY=%LocalAppData%\Programs\Python\Python311\python.exe"
if not defined PY if exist "C:\Program Files\Python312\python.exe" set "PY=C:\Program Files\Python312\python.exe"

if not defined PY (
    echo No se encontro Python en PATH.
    echo Instale Python 3.12 y marque "Add to PATH", o ejecute:
    echo   winget install Python.Python.3.12
    pause
    exit /b 1
)

echo Usando: %PY%
%PY% -m pip install -r requirements.txt -q

for /f "tokens=2 delims==" %%a in ('findstr /B "API_PORT" .env 2^>nul') do set API_PORT=%%a
if not defined API_PORT set API_PORT=8010

echo Iniciando Olnatura Auth en 127.0.0.1:%API_PORT% ...
%PY% -m uvicorn main:app --host 127.0.0.1 --port %API_PORT%
