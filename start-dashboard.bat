@echo off
rem APEX Dashboard launcher - serves this folder locally (read-only) so the dashboard
rem can auto-refresh whenever "DI Ageing-Pends Dashboard.xlsx" is saved.
cd /d "%~dp0"
set PORT=8765

where python >nul 2>nul && (set PY=python) || (set PY=py)

echo Starting APEX Dashboard on http://localhost:%PORT%/ ...
echo Close this window to stop the dashboard server.
start "" "http://localhost:%PORT%/index.html"
%PY% -m http.server %PORT% --bind 127.0.0.1
