# Test-Authentication.ps1
# Test authentication to SharePoint

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl
)

Write-Host "Testing SharePoint Authentication" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host ""

# Load the SharePoint Manager
$scriptPath = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $scriptPath "src\Core\SharePointManager.ps1")

try {
    Write-Host "Creating SharePoint Manager for: $SiteUrl" -ForegroundColor Yellow
    $spManager = [SharePointManager]::new($SiteUrl)
    
    Write-Host "Attempting to connect..." -ForegroundColor Yellow
    if ($spManager.Connect()) {
        Write-Host "✓ Authentication successful!" -ForegroundColor Green
        
        # Try to get site information
        $web = Get-PnPWeb
        Write-Host "✓ Connected to site: $($web.Title)" -ForegroundColor Green
        Write-Host "✓ Site URL: $($web.Url)" -ForegroundColor Green
        
        # Get lists to verify permissions
        Write-Host "`nVerifying permissions..." -ForegroundColor Yellow
        $lists = Get-PnPList | Where-Object { $_.Hidden -eq $false }
        Write-Host "✓ Found $($lists.Count) accessible lists/libraries" -ForegroundColor Green
        
        # Disconnect
        $spManager.Disconnect()
        Write-Host "`n✓ Disconnected successfully" -ForegroundColor Green
        
        Write-Host "`n===================================" -ForegroundColor Cyan
        Write-Host "TEST PASSED" -ForegroundColor Green
        exit 0
    }
    else {
        throw "Authentication failed"
    }
}
catch {
    Write-Host "✗ TEST FAILED: $_" -ForegroundColor Red
    exit 1
}