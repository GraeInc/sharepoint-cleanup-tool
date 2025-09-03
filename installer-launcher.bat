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

REM Check PowerShell version
echo Checking PowerShell version...
powershell -Command "$PSVersionTable.PSVersion.Major"
set /a PSVersion=0
for /f "tokens=*" %%i in ('powershell -Command "$PSVersionTable.PSVersion.Major"') do set PSVersion=%%i
echo PowerShell version: %PSVersion%
if %PSVersion% LSS 7 (
    echo PowerShell 5.1 detected. Installing compatible PnP.PowerShell version...
    goto :install_compatible_module
) else (
    echo PowerShell 7+ detected. Installing latest PnP.PowerShell version...
    goto :check_module
)

:install_compatible_module
echo Installing PnP.PowerShell version 1.12.0 (compatible with PowerShell 5.1)...
powershell -Command "Install-Module -Name PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser -Force -AllowClobber"
if %errorLevel% neq 0 (
    echo ERROR: Failed to install PnP.PowerShell module.
    echo Please check your internet connection and try again.
    pause
    exit /b 1
)
echo PnP.PowerShell module installed successfully.
goto :check_version

:check_module
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
if not exist "sharepoint-cleanup-script.ps1" (
    echo ERROR: sharepoint-cleanup-script.ps1 not found in current directory.
    echo Please ensure all files are extracted to the same folder.
    pause
    exit /b 1
)

if not exist "sharepoint-cleanup-gui.ps1" (
    echo ERROR: sharepoint-cleanup-gui.ps1 not found in current directory.
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
REM Import module and run GUI script in same session
powershell -ExecutionPolicy Bypass -Command "Import-Module PnP.PowerShell; & '.\sharepoint-cleanup-gui.ps1'"
goto :end

:run_cli
echo.
echo Starting Command Line Tool...
echo.
echo Example usage:
echo .\sharepoint-cleanup-script.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -LibraryName "Documents" -ModifiedDate "2024-01-15"
echo.
echo Opening PowerShell window for manual execution...
REM Import module and open PowerShell with module loaded
powershell -NoExit -ExecutionPolicy Bypass -Command "Import-Module PnP.PowerShell; Write-Host 'SharePoint Cleanup Tool - Command Line Mode' -ForegroundColor Green; Write-Host 'PnP.PowerShell module loaded successfully!' -ForegroundColor Green; Write-Host 'Use: .\sharepoint-cleanup-script.ps1 -SiteUrl [URL] -LibraryName [LibName] -ModifiedDate [Date]' -ForegroundColor Yellow; Write-Host 'Example: .\sharepoint-cleanup-script.ps1 -SiteUrl \"https://contoso.sharepoint.com/sites/mysite\" -LibraryName \"Documents\" -ModifiedDate \"2024-01-15\"' -ForegroundColor Cyan"
goto :end

:exit
echo Exiting...
goto :end

:end
echo.
echo Tool execution completed.
pause
exit /b 0