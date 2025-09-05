# Test-GraphSDK.ps1
# Proof of concept for Microsoft Graph SDK as PnP.PowerShell alternative

Write-Host "Microsoft Graph SDK - SharePoint Access Test" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Check if Microsoft.Graph module is installed
$graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Sites
if (-not $graphModule) {
    Write-Host "Microsoft Graph SDK not installed." -ForegroundColor Yellow
    Write-Host "To install, run: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
    Write-Host ""
    $install = Read-Host "Would you like to install it now? (Y/N)"
    if ($install -eq 'Y') {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
        Import-Module Microsoft.Graph.Sites
    } else {
        exit
    }
} else {
    Import-Module Microsoft.Graph.Sites
    Write-Host "Microsoft Graph SDK version: $($graphModule.Version)" -ForegroundColor Green
}

Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Write-Host "This will open a browser for authentication (supports MFA)" -ForegroundColor Cyan
Write-Host ""

try {
    # Connect with delegated permissions - NO APP REGISTRATION NEEDED!
    Connect-MgGraph -Scopes "Sites.Read.All", "Sites.ReadWrite.All" -NoWelcome
    
    $context = Get-MgContext
    Write-Host "Connected successfully!" -ForegroundColor Green
    Write-Host "  Account: $($context.Account)" -ForegroundColor White
    Write-Host "  Tenant: $($context.TenantId)" -ForegroundColor White
    Write-Host ""
    
    # Test SharePoint access
    Write-Host "Enter SharePoint site URL to test:" -ForegroundColor Yellow
    Write-Host "Example: https://contoso.sharepoint.com/sites/TeamSite" -ForegroundColor Gray
    $siteUrl = Read-Host "Site URL"
    
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        $siteUrl = "https://azuregraesecurity.sharepoint.com/sites/GraeInc"
        Write-Host "Using default: $siteUrl" -ForegroundColor Yellow
    }
    
    # Parse the URL to get site ID format
    if ($siteUrl -match "https://([^.]+)\.sharepoint\.com(/sites/[^/]+)?") {
        $hostname = "$($matches[1]).sharepoint.com"
        $sitePath = $matches[2]
        
        if ($sitePath) {
            $siteId = "${hostname}:${sitePath}"
        } else {
            $siteId = $hostname
        }
        
        Write-Host ""
        Write-Host "Accessing site: $siteId" -ForegroundColor Yellow
        
        # Get site information
        $site = Get-MgSite -SiteId $siteId
        Write-Host ""
        Write-Host "Site Information:" -ForegroundColor Cyan
        Write-Host "  Name: $($site.DisplayName)" -ForegroundColor White
        Write-Host "  ID: $($site.Id)" -ForegroundColor White
        Write-Host "  Web URL: $($site.WebUrl)" -ForegroundColor White
        
        # Get document libraries
        Write-Host ""
        Write-Host "Getting document libraries..." -ForegroundColor Yellow
        $lists = Get-MgSiteList -SiteId $site.Id -Filter "baseTemplate eq 101"
        
        Write-Host "Found $($lists.Count) document libraries:" -ForegroundColor Green
        foreach ($list in $lists | Select-Object -First 5) {
            Write-Host "  - $($list.DisplayName)" -ForegroundColor White
            
            # Get item count
            $items = Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id -Top 1
            Write-Host "    Items: Check specific folders for content" -ForegroundColor Gray
        }
        
        Write-Host ""
        Write-Host "✓ Microsoft Graph SDK is working!" -ForegroundColor Green
        Write-Host "✓ Can access SharePoint without app registration!" -ForegroundColor Green
        Write-Host "✓ Supports MFA and modern authentication!" -ForegroundColor Green
        
        Write-Host ""
        Write-Host "Key advantages over PnP.PowerShell:" -ForegroundColor Cyan
        Write-Host "  1. No client ID issues (uses Microsoft's app)" -ForegroundColor White
        Write-Host "  2. Works across all tenants by default" -ForegroundColor White
        Write-Host "  3. Actively maintained and updated" -ForegroundColor White
        Write-Host "  4. Unified API for all Microsoft 365 services" -ForegroundColor White
        
    } else {
        Write-Host "Invalid SharePoint URL format" -ForegroundColor Red
    }
    
} catch {
    Write-Host ""
    Write-Host "Error: $_" -ForegroundColor Red
} finally {
    # Disconnect
    if (Get-MgContext) {
        Disconnect-MgGraph | Out-Null
        Write-Host ""
        Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Yellow
    }
}