@echo off
title Invoice Reminder - Setup
cd /d "%~dp0"
echo.
echo  Invoice Reminder - one-time setup
echo  =================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo  Python is not installed. Install it from https://www.python.org/downloads/
  echo  ^(tick "Add python.exe to PATH" during installation^), then run this setup again.
  echo.
  pause
  exit /b 1
)

echo  [1/3] Installing what the tool needs...
%PY% -m pip install --quiet --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
  echo  Couldn't install the requirements. Check the internet connection and run setup again.
  pause
  exit /b 1
)

echo  [2/3] Shop details...
if not exist .env (
  copy /y .env.example .env >nul
  echo        Notepad will open. Put in the shop name, phone and Paybill,
  echo        then SAVE and CLOSE Notepad to continue.
  start /wait notepad .env
) else (
  echo        .env already exists - keeping your settings.
)
if not exist invoices.xlsx %PY% make_sample_invoices.py

echo  [3/3] Checking invoices every morning at 7:30 AM...
schtasks /Create /TN "Invoice Reminder" /TR "\"%~dp0run_reminders.bat\"" /SC DAILY /ST 07:30 /F >nul
if errorlevel 1 (
  echo  Couldn't create the scheduled task. Right-click setup.bat and choose "Run as administrator".
  pause
  exit /b 1
)
REM Laptops: Windows only runs tasks on mains power by default. Allow battery, and catch up if it was off.
powershell -NoProfile -Command "$s = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable; Set-ScheduledTask -TaskName 'Invoice Reminder' -Settings $s" >nul 2>nul

echo.
echo  Done! Every morning at 7:30 the reminders for overdue invoices
echo  will be waiting in the Outbox folder.
echo.
echo  Next: double-click "check_now.bat" to see today's reminders.
echo.
pause
