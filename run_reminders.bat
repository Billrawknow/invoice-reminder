@echo off
REM Run by Windows Task Scheduler every morning. Writes to automation.log.
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")
%PY% invoice_reminder.py --notify --scheduled
if %errorlevel%==9009 echo %date% %time% ^| ERROR ^| FAILED - scheduled ^| Python was not found>> automation.log
