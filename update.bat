@echo off
rem Daily auto-update (Task Scheduler, 9:00 AM): fetch API data, rebuild settle.js, push to GitHub.
rem Progress shows in this window and is also appended to update.log
cd /d "%~dp0"
echo ==== %date% %time% ==== >> update.log
echo Flipkart data update chal raha hai... (5-10 minute, window band mat kijiye)
powershell -NoProfile -ExecutionPolicy Bypass -Command "& .\fetch.ps1 -Days 30 *>&1 | Tee-Object -FilePath update.log -Append"
powershell -NoProfile -ExecutionPolicy Bypass -Command "& .\build-settle.ps1 *>&1 | Tee-Object -FilePath update.log -Append"
echo GitHub par upload ho raha hai...
powershell -NoProfile -ExecutionPolicy Bypass -Command "& .\fetch-gst.ps1 *>&1 | Tee-Object -FilePath update.log -Append"
powershell -NoProfile -ExecutionPolicy Bypass -Command "& .\fetch-fsn.ps1 *>&1 | Tee-Object -FilePath update.log -Append"
git add data.js settle.js cost.js gst.js fsn.js >> update.log 2>&1
git commit -m "daily data update" >> update.log 2>&1
git push >> update.log 2>&1
if errorlevel 1 (echo GitHub upload FAIL - update.log dekhiye) else (echo Done! Data GitHub par upload ho gaya.)
timeout /t 10
