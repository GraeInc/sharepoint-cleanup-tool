# SharePointCleaner-Simple.ps1
# Simplified launcher for SharePoint Cleanup Tool - Graph SDK Version
# This version avoids the Graph module loading issue

param(
    [switch]$CLI,
    [string]$SiteUrl,
    [string]$LibraryName,
    [datetime]$ModifiedDate,
    [switch]$WhatIf = $true,
    [switch]$ExportResults,
    [string]$ExportPath,
    [switch]$Silent
)

# Check if Microsoft Graph module is installed
Write-Host "Checking for Microsoft Graph SDK..." -ForegroundColor Yellow
$hasGraphAuth = $null -ne (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)
$hasGraphSites = $null -ne (Get-Module -ListAvailable -Name Microsoft.Graph.Sites)

if (-not $hasGraphAuth -or -not $hasGraphSites) {
    Write-Host ""
    Write-Host "Microsoft Graph PowerShell SDK is not installed" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please run Install.ps1 first to install required modules" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Host "Microsoft Graph SDK found" -ForegroundColor Green
Write-Host ""

# Import modules carefully to avoid the parsing error
try {
    # Import without using the problematic internal Graph loading mechanism
    $null = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue
    if (-not $?) {
        Import-Module Microsoft.Graph.Authentication -Force -Global
    }
    
    $null = Get-Command Get-MgSite -ErrorAction SilentlyContinue
    if (-not $?) {
        Import-Module Microsoft.Graph.Sites -Force -Global
    }
    
    Write-Host "Modules loaded successfully" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "Warning: Module loading encountered an issue, but continuing..." -ForegroundColor Yellow
    Write-Host "Error details: $_" -ForegroundColor Gray
    Write-Host ""
}

# Determine mode
if ($CLI -or $PSBoundParameters.Count -gt 1) {
    # CLI mode
    if (-not $SiteUrl -or -not $LibraryName -or -not $ModifiedDate) {
        Write-Host ""
        Write-Host "SharePoint Cleanup Tool - CLI Mode" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Usage:" -ForegroundColor Yellow
        Write-Host "  .\SharePointCleaner-Simple.ps1 -CLI -SiteUrl <url> -LibraryName <name> -ModifiedDate <date> [options]"
        Write-Host ""
        Write-Host "Required Parameters:" -ForegroundColor Yellow
        Write-Host "  -SiteUrl       SharePoint site URL"
        Write-Host "  -LibraryName   Document library name"
        Write-Host "  -ModifiedDate  Filter by modified date"
        Write-Host ""
        Write-Host "Optional Parameters:" -ForegroundColor Yellow
        Write-Host "  -WhatIf        Preview mode (default: `$true)"
        Write-Host "  -ExportResults Export results to CSV"
        Write-Host "  -ExportPath    Path for export file"
        Write-Host "  -Silent        Suppress console output"
        Write-Host ""
        Write-Host "Example:" -ForegroundColor Green
        Write-Host '  .\SharePointCleaner-Simple.ps1 -CLI -SiteUrl "https://contoso.sharepoint.com/sites/TeamSite" `'
        Write-Host '      -LibraryName "Documents" -ModifiedDate "2025-01-15" -WhatIf:$false'
        Write-Host ""
        exit 1
    }
    
    # Run CLI
    $cliPath = Join-Path $PSScriptRoot "src\CLI\CLI.ps1"
    $params = @{
        SiteUrl = $SiteUrl
        LibraryName = $LibraryName
        ModifiedDate = $ModifiedDate
        WhatIf = $WhatIf
    }
    
    if ($ExportResults) { $params.ExportResults = $true }
    if ($ExportPath) { $params.ExportPath = $ExportPath }
    if ($Silent) { $params.Silent = $true }
    
    & $cliPath @params
} else {
    # GUI mode - Run in a new PowerShell instance to avoid module conflicts
    Write-Host "Launching SharePoint Cleanup Tool GUI..." -ForegroundColor Cyan
    Write-Host ""
    
    # Create a temporary launcher script
    $tempScript = [System.IO.Path]::GetTempFileName() + ".ps1"
    
    @'
# Temporary GUI launcher
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set script root
$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $PSScriptRoot) { $PSScriptRoot = Get-Location }

# Import core modules directly
$corePath = Join-Path $PSScriptRoot "src\Core"
$guiPath = Join-Path $PSScriptRoot "src\GUI"

# Load core functions
. (Join-Path $corePath "GraphAuth.ps1")
. (Join-Path $corePath "FolderOps.ps1")
. (Join-Path $corePath "Logger.ps1")

# Load and show GUI
. (Join-Path $guiPath "MainGUI.ps1")
Show-CleanupGUI
'@ | Out-File -FilePath $tempScript -Encoding UTF8
    
    # Update the paths in the temp script
    $scriptContent = Get-Content $tempScript -Raw
    $scriptContent = $scriptContent.Replace('$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', "`$PSScriptRoot = '$PSScriptRoot'")
    $scriptContent | Out-File -FilePath $tempScript -Encoding UTF8 -Force
    
    # Launch GUI in new process
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$tempScript`"" -Wait
    
    # Clean up temp file
    Start-Sleep -Seconds 2
    Remove-Item $tempScript -Force -ErrorAction SilentlyContinue
}