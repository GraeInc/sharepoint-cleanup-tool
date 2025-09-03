@echo off
setlocal

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

REM Check for PowerShell 7
echo Checking for PowerShell 7...
set "PSExe=powershell"

REM First try to find pwsh in PATH (simplest method)
where pwsh >nul 2>&1
if %errorLevel% == 0 (
    set "PSExe=pwsh"
    echo PowerShell 7 found in PATH
) else if exist "C:\Program Files\PowerShell\7\pwsh.exe" (
    set PSExe="C:\Program Files\PowerShell\7\pwsh.exe"
    echo PowerShell 7 found at: C:\Program Files\PowerShell\7\pwsh.exe
) else if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" (
    set PSExe="%ProgramFiles%\PowerShell\7\pwsh.exe"
    echo PowerShell 7 found at: %ProgramFiles%\PowerShell\7\pwsh.exe
) else (
    echo PowerShell 7 not found. Using Windows PowerShell 5.1
    set "PSExe=powershell"
)

REM Check actual PowerShell version
echo Checking PowerShell version...
for /f "tokens=*" %%i in ('%PSExe% -Command "$PSVersionTable.PSVersion.Major"') do set PSVersion=%%i
echo Using PowerShell version: %PSVersion%

REM Check PowerShell execution policy
echo Checking PowerShell execution policy...
%PSExe% -Command "Get-ExecutionPolicy" | findstr /i "restricted" >nul
if %errorLevel% == 0 (
    echo Current execution policy is Restricted. Changing to RemoteSigned...
    %PSExe% -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
    if %errorLevel% neq 0 (
        echo ERROR: Failed to change execution policy. Please run as administrator.
        pause
        exit /b 1
    )
    echo Execution policy updated successfully.
) else (
    echo Execution policy is already permissive.
)
echo.

if %PSVersion% LSS 7 (
    echo WARNING: PowerShell 5.1 detected. The GUI may not work properly.
    echo For best results, please install PowerShell 7 from:
    echo https://github.com/PowerShell/PowerShell/releases
    echo.
    echo Installing compatible PnP.PowerShell version for PowerShell 5.1...
    goto :install_compatible_module
) else (
    echo PowerShell 7+ detected. Installing latest PnP.PowerShell version...
    goto :check_module
)

:install_compatible_module
echo Installing PnP.PowerShell version 1.12.0 (compatible with PowerShell 5.1)...
%PSExe% -Command "Install-Module -Name PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser -Force -AllowClobber"
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
%PSExe% -Command "if (Get-Module -ListAvailable -Name PnP.PowerShell) { exit 0 } else { exit 1 }"
if %errorLevel% == 0 (
    echo PnP.PowerShell module is already installed.
    goto :check_version
) else (
    echo PnP.PowerShell module not found. Installing...
    goto :install_module
)

:install_module
echo Installing PnP.PowerShell module... This may take a few minutes.
%PSExe% -Command "Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber"
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
%PSExe% -Command "Get-Module -ListAvailable -Name PnP.PowerShell | Select-Object Name, Version | Format-Table -AutoSize"
echo.

REM Check if script files exist
if not exist "sharepoint-cleanup-script.ps1" (
    echo ERROR: sharepoint-cleanup-script.ps1 not found in current directory.
    echo Please ensure all files are extracted to the same folder.
    pause
    exit /b 1
)

if not exist "sharepoint-cleanup-simple-gui.ps1" (
    if not exist "sharepoint-cleanup-gui.ps1" (
        echo ERROR: GUI script not found in current directory.
        echo Please ensure all files are extracted to the same folder.
        pause
        exit /b 1
    )
)

echo Setup complete! Choose how to run the tool:
echo.
echo 1. GUI Tool - Windows Interface (Recommended)
echo 2. Command Line Tool (Advanced)
echo 3. Setup Authentication (Troubleshooting)
echo 4. Exit
echo.
set /p choice="Enter your choice (1-4): "

if "%choice%"=="1" goto :run_gui
if "%choice%"=="2" goto :run_cli
if "%choice%"=="3" goto :setup_auth
if "%choice%"=="4" goto :exit
echo Invalid choice. Please enter 1, 2, 3, or 4.
goto :check_version

:run_gui
echo.
echo Starting GUI Tool...
echo.

REM Debug: Show current directory
echo Current directory: %cd%
echo.

REM Re-detect PowerShell for this section
where pwsh >nul 2>&1
if %errorLevel% == 0 (
    set "PSExe=pwsh"
    echo Using PowerShell 7 - pwsh
) else (
    set "PSExe=powershell"
    echo Using Windows PowerShell 5.1
)

echo PowerShell executable: %PSExe%
echo.

REM Launch the GUI - try simple version first (direct launch of CLI)
if exist "sharepoint-cleanup-gui-simple.ps1" (
    echo Found: sharepoint-cleanup-gui-simple.ps1
    echo Launching SharePoint Cleanup GUI - Simple Version...
    echo.
    echo NOTE: This GUI will launch the CLI script in a new window.
    echo You'll handle authentication directly in that window.
    echo.
    echo Executing: Starting simple GUI...
    REM Launch simple GUI in a new window
    start "" %PSExe% -ExecutionPolicy Bypass -File sharepoint-cleanup-gui-simple.ps1
    goto :end
)
if exist "sharepoint-cleanup-gui-wrapper.ps1" (
    echo Found: sharepoint-cleanup-gui-wrapper.ps1
    echo Launching SharePoint Cleanup GUI - CLI Wrapper Version...
    echo.
    echo NOTE: This GUI uses the proven CLI script for all operations.
    echo Browser will open for authentication when you click Run Scan.
    echo.
    echo Executing: Starting GUI wrapper in new window...
    REM Launch wrapper GUI in a new window
    start "" %PSExe% -ExecutionPolicy Bypass -File sharepoint-cleanup-gui-wrapper.ps1
    goto :end
)
if exist "sharepoint-cleanup-gui-final.ps1" (
    echo Found: sharepoint-cleanup-gui-final.ps1
    echo Launching SharePoint Cleanup GUI - Final Version...
    echo.
    echo NOTE: The GUI window will open separately.
    echo Browser will open for authentication when you click Connect.
    echo.
    echo Executing: Starting GUI in new window...
    REM Launch final GUI in a new window
    start "" %PSExe% -ExecutionPolicy Bypass -File sharepoint-cleanup-gui-final.ps1
    goto :end
)
if exist "sharepoint-cleanup-gui-browser.ps1" (
    echo Found: sharepoint-cleanup-gui-browser.ps1
    echo Launching SharePoint Cleanup GUI - Browser Auth Version...
    echo.
    echo NOTE: The GUI window will open separately.
    echo A separate PowerShell window will handle authentication.
    echo.
    echo Executing: Starting GUI in new window...
    REM Launch browser auth GUI in a new window
    start "" %PSExe% -ExecutionPolicy Bypass -File sharepoint-cleanup-gui-browser.ps1
    goto :end
)
if exist "sharepoint-cleanup-gui-async.ps1" (
    echo Found: sharepoint-cleanup-gui-async.ps1
    echo Launching SharePoint Cleanup GUI (Async Version)...
    echo.
    echo NOTE: The GUI window will open separately.
    echo A browser window will open for authentication.
    echo.
    echo Executing: Starting GUI in new window...
    REM Launch async GUI in a new window
    start "" %PSExe% -ExecutionPolicy Bypass -File sharepoint-cleanup-gui-async.ps1
    goto :end
)
if exist "sharepoint-cleanup-gui-working.ps1" (
    echo Found: sharepoint-cleanup-gui-working.ps1
    echo Launching SharePoint Cleanup GUI...
    echo.
    echo NOTE: The GUI window will open separately.
    echo Keep this console window open to see any error messages.
    echo.
    echo Executing: Starting GUI in new window...
    REM Launch GUI in a new window
    start "" %PSExe% -ExecutionPolicy Bypass -File sharepoint-cleanup-gui-working.ps1
    goto :end
)
if exist "run-cleanup.ps1" (
    echo GUI not found, starting interactive cleanup tool...
    echo.
    %PSExe% -ExecutionPolicy Bypass -NoExit -File "run-cleanup.ps1"
    goto :end
)
if exist "sharepoint-cleanup-integrated.ps1" (
    echo Using integrated authentication (no app registration required)...
    echo Running: %PSExe% -ExecutionPolicy Bypass -File "sharepoint-cleanup-integrated.ps1"
    %PSExe% -ExecutionPolicy Bypass -File "sharepoint-cleanup-integrated.ps1"
    if %errorLevel% neq 0 (
        echo.
        echo ERROR: Script execution failed.
        pause
    )
)
if exist "sharepoint-cleanup-simple-gui.ps1" (
    echo Using simplified GUI for better compatibility...
    %PSExe% -ExecutionPolicy Bypass -File "sharepoint-cleanup-simple-gui.ps1"
    if %errorLevel% neq 0 (
        echo.
        echo ERROR: Script execution failed.
        pause
    )
    goto :end
)
REM If no GUI files found - try default GUI
if exist "sharepoint-cleanup-gui.ps1" (
    echo Note: If Windows Defender SmartScreen appears, click "More info" then "Run anyway"
    %PSExe% -ExecutionPolicy Bypass -Command "Import-Module PnP.PowerShell; & '.\sharepoint-cleanup-gui.ps1'"
    if %errorLevel% neq 0 (
        echo.
        echo ERROR: Script execution failed.
        pause
    )
) else (
    echo ERROR: No GUI script found!
    pause
)
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
%PSExe% -NoExit -ExecutionPolicy Bypass -Command "Import-Module PnP.PowerShell; Write-Host 'SharePoint Cleanup Tool - Command Line Mode' -ForegroundColor Green; Write-Host 'PnP.PowerShell module loaded successfully!' -ForegroundColor Green; Write-Host 'Use: .\sharepoint-cleanup-script.ps1 -SiteUrl [URL] -LibraryName [LibName] -ModifiedDate [Date]' -ForegroundColor Yellow; Write-Host 'Example: .\sharepoint-cleanup-script.ps1 -SiteUrl \"https://contoso.sharepoint.com/sites/mysite\" -LibraryName \"Documents\" -ModifiedDate \"2024-01-15\"' -ForegroundColor Cyan"
goto :end

:setup_auth
echo.
echo Starting Authentication Setup...
echo.
if exist "setup-auth.ps1" (
    %PSExe% -ExecutionPolicy Bypass -Command "& '.\setup-auth.ps1'"
) else (
    echo ERROR: setup-auth.ps1 not found.
    echo Please ensure all files are extracted to the same folder.
    pause
)
goto :end

:exit
echo Exiting...
goto :end

:end
echo.
echo Tool execution completed.
pause
exit /b 0