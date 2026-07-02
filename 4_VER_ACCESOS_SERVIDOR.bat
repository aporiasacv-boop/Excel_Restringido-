@echo off
title Ver ultimos accesos (auditoria en servidor)
cd /d "%~dp0api"
set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py -3.12"
if not defined PY where py >nul 2>&1 && set "PY=py -3"
%PY% -c "import database as db; db.init_db(); rows=db.list_audit(30); print('Ultimos accesos en PC de TI (SQLite):'); print(''); [print(f\"{r['created_at']}  {r['username']}  {r.get('workbook') or '-'}  {r['action']}\") for r in rows] if rows else print('(sin registros)')"
echo.
pause
