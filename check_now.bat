@echo off
title Invoice Reminder - check now
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")
%PY% invoice_reminder.py
echo.
if exist Outbox start "" explorer Outbox
pause
