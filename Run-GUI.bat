@echo off
echo Starting SharePoint Cleanup Tool GUI...
echo.

REM Check for PowerShell 7 first
where pwsh >nul 2>&1
if %errorLevel% == 0 (
    pwsh -ExecutionPolicy Bypass -File ".\SharePointCleanup.ps1" -GUI
) else (
    powershell -ExecutionPolicy Bypass -File ".\SharePointCleanup.ps1" -GUI
)

if %errorLevel% neq 0 (
    echo.
    echo ERROR: Failed to launch GUI
    pause
)