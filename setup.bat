@echo off
title Invoice Reminder - Setup
cd /d "%~dp0"
echo.
echo  Invoice Reminder - setup
echo  ========================
echo.

if exist "InvoiceReminder.exe" (
  set APP="%~dp0InvoiceReminder.exe"
  goto :settings
)

REM Developer copy (no .exe): use Python.
set PY=python
where py >nul 2>nul && set PY=py -3
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo  Python is not installed. Install it from https://www.python.org/downloads/
  echo  ^(tick "Add python.exe to PATH"^), then run setup again.
  pause
  exit /b 1
)
echo  Installing what the tool needs...
%PY% -m pip install --quiet --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
  echo  Couldn't install the requirements. Check the internet connection and run setup again.
  pause
  exit /b 1
)
set APP=%PY% "%~dp0invoice_reminder.py"

:settings
if not exist config.env copy /y config.env.example config.env >nul
echo  [1/2] Your settings
echo        Notepad will open. Fill in your shop name, phone and Paybill,
echo        then SAVE (Ctrl+S) and CLOSE Notepad to continue.
:edit
start /wait notepad config.env
%APP% --check
if errorlevel 1 (
  echo  Notepad will open again so you can fix the items above.
  pause
  goto :edit
)

echo  [2/2] Turning on the daily check...
%APP% --install-schedule
if errorlevel 1 (
  pause
  exit /b 1
)
echo.
echo  All set. Double-click run.bat any time to see today's reminders.
echo.
pause
