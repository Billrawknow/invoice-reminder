@echo off
schtasks /Delete /TN "Invoice Reminder" /F
echo The morning check has been turned off. Run setup.bat to turn it back on.
pause
