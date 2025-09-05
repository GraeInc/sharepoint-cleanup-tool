# Install.ps1
# Installer for SharePoint Cleanup Tool - Graph SDK Version

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " SharePoint Cleanup Tool - Graph SDK Installer" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host ""

# Check PowerShell version
$psVersion = $PSVersionTable.PSVersion
Write-Host "PowerShell Version: $($psVersion.Major).$($psVersion.Minor)" -ForegroundColor Yellow

if ($psVersion.Major -lt 5) {
    Write-Host "ERROR: PowerShell 5.1 or later is required" -ForegroundColor Red
    Write-Host "Please upgrade PowerShell and try again" -ForegroundColor Yellow
    exit 1
}

Write-Host "PowerShell version check: OK" -ForegroundColor Green
Write-Host ""

# Check execution policy
$executionPolicy = Get-ExecutionPolicy -Scope CurrentUser
Write-Host "Execution Policy (CurrentUser): $executionPolicy" -ForegroundColor Yellow

if ($executionPolicy -eq "Restricted" -or $executionPolicy -eq "AllSigned") {
    Write-Host "Setting execution policy to RemoteSigned for current user..." -ForegroundColor Yellow
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Host "Execution policy updated successfully" -ForegroundColor Green
    } catch {
        Write-Host "WARNING: Could not update execution policy" -ForegroundColor Yellow
        Write-Host "You may need to run: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Yellow
    }
} else {
    Write-Host "Execution policy check: OK" -ForegroundColor Green
}

Write-Host ""

# Check if Microsoft Graph module is installed
Write-Host "Checking Microsoft Graph PowerShell SDK..." -ForegroundColor Yellow

$graphModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Sites"
)

$modulesNeeded = @()

foreach ($moduleName in $graphModules) {
    $module = Get-Module -ListAvailable -Name $moduleName
    if ($module) {
        Write-Host "  OK: $moduleName (v$($module.Version))" -ForegroundColor Green
    } else {
        Write-Host "  Missing: $moduleName not found" -ForegroundColor Yellow
        $modulesNeeded += $moduleName
    }
}

Write-Host ""

# Install missing modules
if ($modulesNeeded.Count -gt 0) {
    Write-Host "Installing Microsoft Graph PowerShell SDK..." -ForegroundColor Yellow
    Write-Host "This may take a few minutes..." -ForegroundColor Gray
    
    try {
        # Install the main module which includes all sub-modules
        Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
        Write-Host ""
        Write-Host "Microsoft Graph SDK installed successfully!" -ForegroundColor Green
    } catch {
        Write-Host ""
        Write-Host "ERROR: Failed to install Microsoft Graph SDK" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "Please try manual installation:" -ForegroundColor Yellow
        Write-Host "  Install-Module -Name Microsoft.Graph -Scope CurrentUser" -ForegroundColor Cyan
        exit 1
    }
} else {
    Write-Host "All required modules are already installed" -ForegroundColor Green
}

Write-Host ""

# Create necessary directories
Write-Host "Setting up directory structure..." -ForegroundColor Yellow

$dirs = @("logs", "config")
foreach ($dir in $dirs) {
    $path = Join-Path $PSScriptRoot $dir
    if (-not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
        Write-Host "  Created: $dir\" -ForegroundColor Green
    } else {
        Write-Host "  Exists: $dir\" -ForegroundColor Gray
    }
}

Write-Host ""

# Create default configuration
$configPath = Join-Path $PSScriptRoot "config\settings.json"
if (-not (Test-Path $configPath)) {
    $defaultConfig = @{
        DefaultLibrary = "Documents"
        PreviewMode = $true
        MaxBatchSize = 100
        EnableLogging = $true
        LogRetentionDays = 30
        RecentSites = @()
    } | ConvertTo-Json -Depth 3
    
    $defaultConfig | Out-File -FilePath $configPath -Encoding UTF8
    Write-Host "Created default configuration" -ForegroundColor Green
} else {
    Write-Host "Configuration file already exists" -ForegroundColor Gray
}

Write-Host ""

# Create batch launcher
$launcherPath = Join-Path $PSScriptRoot "Launch-GUI.bat"
@'
@echo off
echo Launching SharePoint Cleanup Tool...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0SharePointCleaner.ps1"
pause
'@ | Out-File -FilePath $launcherPath -Encoding ASCII

Write-Host "Created launcher: Launch-GUI.bat" -ForegroundColor Green

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
Write-Host " Installation Complete!" -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host ""
Write-Host "To launch the GUI:" -ForegroundColor Yellow
Write-Host "  - Double-click 'Launch-GUI.bat'" -ForegroundColor Cyan
Write-Host "  - Or run: .\SharePointCleaner.ps1" -ForegroundColor Cyan
Write-Host ""
Write-Host "For CLI mode:" -ForegroundColor Yellow
Write-Host "  .\SharePointCleaner.ps1 -CLI -SiteUrl [url] -LibraryName [library] -ModifiedDate [date]" -ForegroundColor Cyan
Write-Host ""
Write-Host "Note: This tool uses Microsoft Graph SDK and does NOT require" -ForegroundColor Yellow
Write-Host "      any Azure app registration. Authentication is handled" -ForegroundColor Yellow
Write-Host "      automatically through the Graph SDK." -ForegroundColor Yellow
Write-Host ""