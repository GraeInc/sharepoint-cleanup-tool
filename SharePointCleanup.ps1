# SharePointCleanup.ps1
# Main launcher for SharePoint Cleanup Tool

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$GUI,
    
    [Parameter()]
    [switch]$CLI,
    
    [Parameter()]
    [string]$SiteUrl,
    
    [Parameter()]
    [string]$LibraryName,
    
    [Parameter()]
    [datetime]$ModifiedDate,
    
    [Parameter()]
    [switch]$WhatIf = $true
)

$ErrorActionPreference = 'Stop'

# Check if running as administrator (recommended but not required)
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Warning "Not running as administrator. Some features may be limited."
}

# Check for PnP.PowerShell module
Write-Host "Checking for PnP.PowerShell module..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
    
    try {
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        Write-Host "PnP.PowerShell module installed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install PnP.PowerShell module: $_"
        Write-Host "Please install manually using: Install-Module -Name PnP.PowerShell -Scope CurrentUser" -ForegroundColor Red
        exit 1
    }
}

# If no parameters provided or GUI flag set, launch GUI
if ($GUI -or (-not $CLI -and -not $PSBoundParameters.ContainsKey('SiteUrl'))) {
    Write-Host "Launching GUI mode..." -ForegroundColor Green
    $guiPath = Join-Path $PSScriptRoot "SharePointCleanup-GUI.ps1"
    
    if (Test-Path $guiPath) {
        & $guiPath
    }
    else {
        Write-Error "GUI script not found at: $guiPath"
        exit 1
    }
}
# CLI mode
elseif ($CLI -or $PSBoundParameters.ContainsKey('SiteUrl')) {
    if (-not $SiteUrl -or -not $LibraryName -or -not $ModifiedDate) {
        Write-Error "CLI mode requires -SiteUrl, -LibraryName, and -ModifiedDate parameters"
        Write-Host "`nUsage:" -ForegroundColor Yellow
        Write-Host "  .\SharePointCleanup.ps1 -CLI -SiteUrl <url> -LibraryName <name> -ModifiedDate <date> [-WhatIf]"
        Write-Host "`nExample:" -ForegroundColor Yellow
        Write-Host '  .\SharePointCleanup.ps1 -CLI -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -LibraryName "Documents" -ModifiedDate "2024-01-15"'
        exit 1
    }
    
    Write-Host "Launching CLI mode..." -ForegroundColor Green
    $cliPath = Join-Path $PSScriptRoot "SharePointCleanup-CLI.ps1"
    
    if (Test-Path $cliPath) {
        & $cliPath -SiteUrl $SiteUrl -LibraryName $LibraryName -ModifiedDate $ModifiedDate -WhatIf:$WhatIf
    }
    else {
        Write-Error "CLI script not found at: $cliPath"
        exit 1
    }
}
else {
    # Show usage
    Write-Host "SharePoint Cleanup Tool v2.0" -ForegroundColor Cyan
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  GUI Mode (default):"
    Write-Host "    .\SharePointCleanup.ps1"
    Write-Host "    .\SharePointCleanup.ps1 -GUI"
    Write-Host ""
    Write-Host "  CLI Mode:"
    Write-Host "    .\SharePointCleanup.ps1 -CLI -SiteUrl <url> -LibraryName <name> -ModifiedDate <date> [-WhatIf]"
    Write-Host ""
    Write-Host "Parameters:" -ForegroundColor Yellow
    Write-Host "  -GUI           : Launch graphical user interface"
    Write-Host "  -CLI           : Use command-line interface"
    Write-Host "  -SiteUrl       : SharePoint site URL"
    Write-Host "  -LibraryName   : Document library name"
    Write-Host "  -ModifiedDate  : Date when folders were modified"
    Write-Host "  -WhatIf        : Preview mode (default: `$true)"
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host '  .\SharePointCleanup.ps1'
    Write-Host '  .\SharePointCleanup.ps1 -CLI -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -LibraryName "Documents" -ModifiedDate "2024-01-15"'
    Write-Host '  .\SharePointCleanup.ps1 -CLI -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -LibraryName "Documents" -ModifiedDate "2024-01-15" -WhatIf:$false'
}