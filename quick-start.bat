@echo off
echo ========================================
echo SharePoint Cleanup Tool - Quick Start
echo ========================================
echo.
echo This will launch the GUI tool directly.
echo For full setup options, use install-and-run.bat
echo.

REM Check if PnP.PowerShell is available
powershell -Command "if (Get-Module -ListAvailable -Name PnP.PowerShell) { exit 0 } else { exit 1 }"
if %errorLevel% neq 0 (
    echo PnP.PowerShell module not found!
    echo Please run install-and-run.bat first to install dependencies.
    echo.
    pause
    exit /b 1
)

REM Check if GUI script exists
if not exist "SharePoint-Cleanup-GUI.ps1" (
    echo SharePoint-Cleanup-GUI.ps1 not found!
    echo Please ensure all files are in the same folder.
    echo.
    pause
    exit /b 1
)

echo Launching GUI Tool...
powershell -ExecutionPolicy Bypass -File "SharePoint-Cleanup-GUI.ps1"

echo.
echo Tool closed. Press any key to exit.
pause >nul