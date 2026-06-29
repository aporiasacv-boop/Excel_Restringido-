@echo off
cd /d "%~dp0"
if not exist .env copy .env.example .env
python -m pip install -r requirements.txt -q
for /f "tokens=2 delims==" %%a in ('findstr /B "API_PORT" .env 2^>nul') do set API_PORT=%%a
if not defined API_PORT set API_PORT=8010
echo Iniciando Olnatura Auth en 127.0.0.1:%API_PORT% ...
python -m uvicorn main:app --host 127.0.0.1 --port %API_PORT%
