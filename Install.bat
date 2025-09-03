@echo off
setlocal

echo ============================================
echo SharePoint Cleanup Tool v2.0 - Installer
echo ============================================
echo.

REM Check for PowerShell
where powershell >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: PowerShell not found!
    echo Please ensure PowerShell is installed and in PATH.
    pause
    exit /b 1
)

REM Check for PowerShell 7 (optional but recommended)
where pwsh >nul 2>&1
if %errorLevel% == 0 (
    set "PSEXE=pwsh"
    echo Found PowerShell 7
) else (
    set "PSEXE=powershell"
    echo Using Windows PowerShell
)

echo.
echo Checking execution policy...
%PSEXE% -Command "if ((Get-ExecutionPolicy -Scope CurrentUser) -eq 'Restricted') { exit 1 } else { exit 0 }"
if %errorLevel% neq 0 (
    echo Setting execution policy...
    %PSEXE% -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
)

echo.
echo Installing PnP.PowerShell module...
%PSEXE% -Command "if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) { Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber; Write-Host 'Module installed successfully!' -ForegroundColor Green } else { Write-Host 'Module already installed!' -ForegroundColor Green }"

echo.
echo ============================================
echo Installation complete!
echo ============================================
echo.
echo You can now run the tool using:
echo   GUI Mode: PowerShell .\SharePointCleanup.ps1
echo   CLI Mode: PowerShell .\SharePointCleanup.ps1 -CLI -SiteUrl [url] -LibraryName [name] -ModifiedDate [date]
echo.
echo For convenience, you can also double-click Run-GUI.bat
echo.
pause