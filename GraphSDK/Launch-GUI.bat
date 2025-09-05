@echo off
echo Launching SharePoint Cleanup Tool...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0SharePointCleaner.ps1"
pause
