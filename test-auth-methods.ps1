# Test SharePoint Authentication Methods
param(
    [string]$SiteUrl = "https://yourtenant.sharepoint.com/sites/yoursite"
)

Write-Host "Testing SharePoint Authentication Methods" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host ""

if (-not $SiteUrl -or $SiteUrl -eq "https://yourtenant.sharepoint.com/sites/yoursite") {
    $SiteUrl = Read-Host "Enter your SharePoint site URL"
}

Write-Host "Testing site: $SiteUrl" -ForegroundColor Yellow
Write-Host ""

# Test 1: Interactive
Write-Host "Test 1: Interactive Authentication" -ForegroundColor Green
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive
    $web = Get-PnPWeb
    Write-Host "✓ SUCCESS: Connected to $($web.Title)" -ForegroundColor Green
    Disconnect-PnPOnline
}
catch {
    Write-Host "✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# Test 2: UseWebLogin
Write-Host "Test 2: UseWebLogin Authentication" -ForegroundColor Green
try {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    $web = Get-PnPWeb
    Write-Host "✓ SUCCESS: Connected to $($web.Title)" -ForegroundColor Green
    Disconnect-PnPOnline
}
catch {
    Write-Host "✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# Test 3: LaunchBrowser
Write-Host "Test 3: LaunchBrowser Authentication" -ForegroundColor Green
try {
    Connect-PnPOnline -Url $SiteUrl -LaunchBrowser
    $web = Get-PnPWeb
    Write-Host "✓ SUCCESS: Connected to $($web.Title)" -ForegroundColor Green
    Disconnect-PnPOnline
}
catch {
    Write-Host "✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# Test 4: DeviceLogin
Write-Host "Test 4: DeviceLogin Authentication" -ForegroundColor Green
try {
    Connect-PnPOnline -Url $SiteUrl -DeviceLogin
    $web = Get-PnPWeb
    Write-Host "✓ SUCCESS: Connected to $($web.Title)" -ForegroundColor Green
    Disconnect-PnPOnline
}
catch {
    Write-Host "✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Testing complete!" -ForegroundColor Cyan
Read-Host "Press Enter to exit"