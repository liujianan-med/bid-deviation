@echo off
rem Start the deviation-table web service. Browser: http://127.0.0.1:8000
cd /d %~dp0
chcp 65001 >nul
echo Starting bid-deviation web service...
echo Open in browser: http://127.0.0.1:8000
echo Close this window to stop the service.
start "" http://127.0.0.1:8000
.venv\Scripts\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port 8000
pause
