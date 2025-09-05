@echo off
echo ============================================================
echo  SharePoint Cleanup Tool - Graph SDK Version
echo ============================================================
echo.
echo Checking for PowerShell 7...

where pwsh >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: PowerShell 7 is not installed!
    echo.
    echo The Microsoft Graph module has issues with Windows PowerShell 5.1
    echo Please install PowerShell 7 from: https://aka.ms/powershell
    echo.
    echo Or try using the Direct version:
    echo   SharePointCleaner-Direct.ps1
    echo.
    pause
    exit /b 1
)

echo PowerShell 7 found - launching tool...
echo.
pwsh.exe -ExecutionPolicy Bypass -File "%~dp0SharePointCleaner.ps1"
pause