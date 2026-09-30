@echo off
rem Daily auto-update (Task Scheduler, 9:00 AM): fetch API data, rebuild settle.js; log to update.log
cd /d "%~dp0"
echo ==== %date% %time% ==== >> update.log
powershell -NoProfile -ExecutionPolicy Bypass -File fetch.ps1 -Days 30 >> update.log 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File build-settle.ps1 >> update.log 2>&1
git add data.js settle.js >> update.log 2>&1
git commit -m "daily data update" >> update.log 2>&1
git push >> update.log 2>&1
