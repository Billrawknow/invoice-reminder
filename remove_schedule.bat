@echo off
cd /d "%~dp0"
if exist "InvoiceReminder.exe" (
  "%~dp0InvoiceReminder.exe" --remove-schedule
) else (
  where py >nul 2>nul
  if errorlevel 1 (python "%~dp0invoice_reminder.py" --remove-schedule) else (py -3 "%~dp0invoice_reminder.py" --remove-schedule)
)
pause
