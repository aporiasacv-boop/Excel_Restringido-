@echo off
cd /d "%~dp0"

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py -3.12"
if not defined PY where py >nul 2>&1 && set "PY=py -3"
if not defined PY if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PY=%LocalAppData%\Programs\Python\Python312\python.exe"

if not defined PY (
    echo Python no encontrado.
    exit /b 1
)

if "%~2"=="" (
    echo Uso: crear_usuario.bat ^<usuario^> ^<contrasena^>
    exit /b 1
)

%PY% crear_usuario.py %1 %2
