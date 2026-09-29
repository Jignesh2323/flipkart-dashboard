@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File fetch.ps1 -Days 30
powershell -NoProfile -ExecutionPolicy Bypass -File build-settle.ps1
start "" dashboard.html
pause
