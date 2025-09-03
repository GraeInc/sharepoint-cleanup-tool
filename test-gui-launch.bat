@echo off
echo Testing GUI Launch...
echo.

REM Check for PowerShell 7
where pwsh >nul 2>&1
if %errorLevel% == 0 (
    echo Found PowerShell 7 (pwsh)
    echo Launching GUI with PowerShell 7...
    pwsh -ExecutionPolicy Bypass -NoExit -Command "& '.\sharepoint-cleanup-gui-working.ps1'"
) else (
    echo PowerShell 7 not found, using Windows PowerShell
    echo Launching GUI with Windows PowerShell...
    powershell -ExecutionPolicy Bypass -NoExit -Command "& '.\sharepoint-cleanup-gui-working.ps1'"
)

pause