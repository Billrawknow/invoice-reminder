@echo off
title Invoice Reminder
cd /d "%~dp0"
if exist "InvoiceReminder.exe" (
  "%~dp0InvoiceReminder.exe" --open
) else (
  where py >nul 2>nul
  if errorlevel 1 (python "%~dp0invoice_reminder.py" --open) else (py -3 "%~dp0invoice_reminder.py" --open)
)
echo.
pause
