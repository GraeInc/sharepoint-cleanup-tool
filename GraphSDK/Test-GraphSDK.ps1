# Test-GraphSDK.ps1
# Test script for SharePoint Cleanup Tool - Graph SDK Version

Write-Host ""
Write-Host "SharePoint Cleanup Tool - Graph SDK Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Test 1: Check modules
Write-Host "Test 1: Checking Microsoft Graph modules..." -ForegroundColor Yellow
$modules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Sites")
$allFound = $true

foreach ($module in $modules) {
    $m = Get-Module -ListAvailable -Name $module
    if ($m) {
        Write-Host "  ✓ $module found (v$($m.Version))" -ForegroundColor Green
    } else {
        Write-Host "  ✗ $module not found" -ForegroundColor Red
        $allFound = $false
    }
}

if (-not $allFound) {
    Write-Host ""
    Write-Host "Please run Install.ps1 first" -ForegroundColor Yellow
    exit 1
}

Write-Host ""

# Test 2: Import modules
Write-Host "Test 2: Loading modules..." -ForegroundColor Yellow
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Sites -ErrorAction Stop
    Write-Host "  ✓ Modules loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Failed to load modules: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Test 3: Test authentication
Write-Host "Test 3: Testing Microsoft Graph authentication..." -ForegroundColor Yellow
Write-Host "  A browser window will open for authentication" -ForegroundColor Cyan
Write-Host "  Please sign in with your Microsoft 365 account" -ForegroundColor Cyan
Write-Host ""

try {
    Connect-MgGraph -Scopes "Sites.ReadWrite.All" -NoWelcome -ErrorAction Stop
    $context = Get-MgContext
    
    if ($context) {
        Write-Host "  ✓ Authentication successful!" -ForegroundColor Green
        Write-Host "    Account: $($context.Account)" -ForegroundColor Gray
        Write-Host "    Tenant: $($context.TenantId)" -ForegroundColor Gray
        Write-Host "    Scopes: $($context.Scopes -join ', ')" -ForegroundColor Gray
    } else {
        Write-Host "  ✗ Authentication failed" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "  ✗ Authentication error: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Test 4: SharePoint access
Write-Host "Test 4: Testing SharePoint access..." -ForegroundColor Yellow
Write-Host "Enter a SharePoint site URL to test (or press Enter for example):" -ForegroundColor Cyan
$testUrl = Read-Host "URL"

if (-not $testUrl) {
    Write-Host "  No URL provided, skipping SharePoint test" -ForegroundColor Gray
} else {
    try {
        # Convert URL to Graph Site ID
        if ($testUrl -match "https://([^.]+)\.sharepoint\.com(/sites/[^/]+)?") {
            $hostname = "$($matches[1]).sharepoint.com"
            $sitePath = $matches[2]
            $siteId = if ($sitePath) { "${hostname}:${sitePath}" } else { $hostname }
            
            Write-Host "  Testing site: $siteId" -ForegroundColor Gray
            
            $site = Get-MgSite -SiteId $siteId -ErrorAction Stop
            Write-Host "  ✓ SharePoint access successful!" -ForegroundColor Green
            Write-Host "    Site Name: $($site.DisplayName)" -ForegroundColor Gray
            Write-Host "    Web URL: $($site.WebUrl)" -ForegroundColor Gray
            
            # Try to get libraries
            $lists = Get-MgSiteList -SiteId $site.Id -Filter "baseTemplate eq 101" -Top 5
            Write-Host "    Document Libraries: $($lists.Count) found" -ForegroundColor Gray
        } else {
            Write-Host "  ✗ Invalid SharePoint URL format" -ForegroundColor Red
        }
    } catch {
        Write-Host "  ✗ SharePoint access failed: $_" -ForegroundColor Red
        Write-Host "    This might be a permissions issue or invalid URL" -ForegroundColor Yellow
    }
}

# Cleanup
Write-Host ""
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
Disconnect-MgGraph | Out-Null
Write-Host "  ✓ Disconnected" -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " All tests completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "The tool is ready to use:" -ForegroundColor Yellow
Write-Host "  GUI Mode: .\SharePointCleaner.ps1" -ForegroundColor Cyan
Write-Host "  CLI Mode: .\SharePointCleaner.ps1 -CLI -SiteUrl <url> -LibraryName <lib> -ModifiedDate <date>" -ForegroundColor Cyan
Write-Host ""