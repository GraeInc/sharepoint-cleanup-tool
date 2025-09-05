# SharePointCleaner.ps1
# Main launcher for SharePoint Cleanup Tool - Graph SDK Version

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
$graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Sites
if (-not $graphModule) {
    Write-Host ""
    Write-Host "Microsoft Graph PowerShell SDK is not installed" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please run Install.ps1 first to install required modules" -ForegroundColor Yellow
    Write-Host "Or manually install with: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host ""
    exit 1
}

# Import Microsoft Graph modules
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Sites -ErrorAction Stop
} catch {
    Write-Host "Failed to load Microsoft Graph modules: $_" -ForegroundColor Red
    exit 1
}

# Determine mode
if ($CLI -or $PSBoundParameters.Count -gt 0) {
    # CLI mode
    if (-not $SiteUrl -or -not $LibraryName -or -not $ModifiedDate) {
        Write-Host ""
        Write-Host "SharePoint Cleanup Tool - CLI Mode" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Usage:" -ForegroundColor Yellow
        Write-Host "  .\SharePointCleaner.ps1 -CLI -SiteUrl <url> -LibraryName <name> -ModifiedDate <date> [options]"
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
        Write-Host '  .\SharePointCleaner.ps1 -CLI -SiteUrl "https://contoso.sharepoint.com/sites/TeamSite" `'
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
    # GUI mode
    Write-Host ""
    Write-Host "Launching SharePoint Cleanup Tool GUI..." -ForegroundColor Cyan
    
    $guiPath = Join-Path $PSScriptRoot "src\GUI\MainGUI.ps1"
    . $guiPath
    
    Show-CleanupGUI
}