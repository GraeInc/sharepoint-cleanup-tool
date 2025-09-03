@echo off
echo Testing GUI Launch
echo.

where pwsh >nul 2>&1
if %errorLevel% == 0 (
    echo Found PowerShell 7
    pwsh -ExecutionPolicy Bypass -File sharepoint-cleanup-gui-browser.ps1
) else (
    echo Using Windows PowerShell
    powershell -ExecutionPolicy Bypass -File sharepoint-cleanup-gui-browser.ps1
)

pause
