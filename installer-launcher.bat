@echo off
setlocal enabledelayedexpansion

echo ========================================
echo SharePoint Empty Folder Cleanup Tool
echo ========================================
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running as administrator...
) else (
    echo WARNING: Not running as administrator. Some installations may fail.
    echo Right-click this file and select "Run as administrator" for best results.
    echo.
    pause
)

REM Check PowerShell execution policy
echo Checking PowerShell execution policy...
powershell -Command "Get-ExecutionPolicy" | findstr /i "restricted" >nul
if %errorLevel% == 0 (
    echo Current execution policy is Restricted. Changing to RemoteSigned...
    powershell -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
    if !errorLevel! neq 0 (
        echo ERROR: Failed to change execution policy. Please run as administrator.
        pause
        exit /b 1
    )
    echo Execution policy updated successfully.
) else (
    echo Execution policy is already permissive.
)
echo.

REM Check if PnP.PowerShell module is installed
echo Checking for PnP.PowerShell module...
powershell -Command "if (Get-Module -ListAvailable -Name PnP.PowerShell) { exit 0 } else { exit 1 }"
if %errorLevel% == 0 (
    echo PnP.PowerShell module is already installed.
    goto :check_version
) else (
    echo PnP.PowerShell module not found. Installing...
    goto :install_module
)

:install_module
echo Installing PnP.PowerShell module... This may take a few minutes.
powershell -Command "Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber"
if %errorLevel% neq 0 (
    echo ERROR: Failed to install PnP.PowerShell module.
    echo Please check your internet connection and try again.
    pause
    exit /b 1
)
echo PnP.PowerShell module installed successfully.
echo.

:check_version
echo Checking module version...
powershell -Command "Get-Module -ListAvailable -Name PnP.PowerShell | Select-Object Name, Version | Format-Table -AutoSize"
echo.

REM Check if script files exist
if not exist "SharePoint-Cleanup.ps1" (
    echo ERROR: SharePoint-Cleanup.ps1 not found in current directory.
    echo Please ensure all files are extracted to the same folder.
    pause
    exit /b 1
)

if not exist "SharePoint-Cleanup-GUI.ps1" (
    echo ERROR: SharePoint-Cleanup-GUI.ps1 not found in current directory.
    echo Please ensure all files are extracted to the same folder.
    pause
    exit /b 1
)

echo Setup complete! Choose how to run the tool:
echo.
echo 1. GUI Tool (Recommended for first-time users)
echo 2. Command Line Tool
echo 3. Exit
echo.
set /p choice="Enter your choice (1-3): "

if "%choice%"=="1" goto :run_gui
if "%choice%"=="2" goto :run_cli
if "%choice%"=="3" goto :exit
echo Invalid choice. Please enter 1, 2, or 3.
goto :check_version

:run_gui
echo.
echo Starting GUI Tool...
echo Note: If Windows Defender SmartScreen appears, click "More info" then "Run anyway"
echo.
powershell -ExecutionPolicy Bypass -File "SharePoint-Cleanup-GUI.ps1"
goto :end

:run_cli
echo.
echo Starting Command Line Tool...
echo.
echo Example usage:
echo .\SharePoint-Cleanup.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -LibraryName "Documents" -ModifiedDate "2024-01-15"
echo.
echo Opening PowerShell window for manual execution...
powershell -NoExit -ExecutionPolicy Bypass -Command "Write-Host 'SharePoint Cleanup Tool - Command Line Mode' -ForegroundColor Green; Write-Host 'Use: .\SharePoint-Cleanup.ps1 -SiteUrl [URL] -LibraryName [LibName] -ModifiedDate [Date]' -ForegroundColor Yellow; Write-Host 'Example: .\SharePoint-Cleanup.ps1 -SiteUrl \"https://contoso.sharepoint.com/sites/mysite\" -LibraryName \"Documents\" -ModifiedDate \"2024-01-15\"' -ForegroundColor Cyan"
goto :end

:exit
echo Exiting...
goto :end

:end
echo.
echo Tool execution completed.
pause
exit /b 0