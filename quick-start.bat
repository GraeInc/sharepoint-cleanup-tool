@echo off
echo ========================================
echo SharePoint Cleanup Tool - Quick Start
echo ========================================
echo.
echo This will launch the GUI tool directly.
echo For full setup options, use installer-launcher.bat
echo.

REM Check PowerShell version and import appropriate module
echo Checking PowerShell version and importing PnP.PowerShell...
powershell -Command "$PSVersion = $PSVersionTable.PSVersion.Major; if ($PSVersion -lt 7) { Import-Module PnP.PowerShell -RequiredVersion 1.12.0 } else { Import-Module PnP.PowerShell }; if (Get-Module PnP.PowerShell) { exit 0 } else { exit 1 }"
if %errorLevel% neq 0 (
    echo PnP.PowerShell module not found or failed to import!
    echo Please run installer-launcher.bat first to install dependencies.
    echo.
    pause
    exit /b 1
)

REM Check if GUI script exists
if not exist "sharepoint-cleanup-gui.ps1" (
    echo sharepoint-cleanup-gui.ps1 not found!
    echo Please ensure all files are in the same folder.
    echo.
    pause
    exit /b 1
)

echo Launching GUI Tool...
REM Import module and run GUI script in same session
powershell -ExecutionPolicy Bypass -Command "$PSVersion = $PSVersionTable.PSVersion.Major; if ($PSVersion -lt 7) { Import-Module PnP.PowerShell -RequiredVersion 1.12.0 } else { Import-Module PnP.PowerShell }; & '.\sharepoint-cleanup-gui.ps1'"

echo.
echo Tool closed. Press any key to exit.
pause >nul