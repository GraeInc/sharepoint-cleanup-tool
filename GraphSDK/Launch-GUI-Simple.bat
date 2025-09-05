@echo off
echo Launching SharePoint Cleanup Tool (Simplified)...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0SharePointCleaner-Simple.ps1"
pause