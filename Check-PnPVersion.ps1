# Check-PnPVersion.ps1
# Check which PnP.PowerShell version is installed and what auth methods are available

Write-Host "Checking PnP.PowerShell Module Configuration" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Get module info
$modules = Get-Module -ListAvailable -Name PnP.PowerShell
Write-Host "Installed PnP.PowerShell versions:" -ForegroundColor Yellow
foreach ($module in $modules) {
    Write-Host "  Version: $($module.Version)" -ForegroundColor White
}
Write-Host ""

# Check available parameters for Connect-PnPOnline
Write-Host "Checking available authentication parameters..." -ForegroundColor Yellow
$command = Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue

if ($command) {
    $params = $command.Parameters.Keys | Where-Object { 
        $_ -match "Interactive|DeviceLogin|UseWebLogin|Credentials|ClientId|ClientSecret|LaunchBrowser|Tenant|ManagedIdentity"
    }
    
    Write-Host "Available authentication parameters:" -ForegroundColor Green
    foreach ($param in $params | Sort-Object) {
        Write-Host "  -$param" -ForegroundColor White
    }
}
else {
    Write-Host "Connect-PnPOnline command not found!" -ForegroundColor Red
}

Write-Host ""
Write-Host "Testing authentication methods..." -ForegroundColor Yellow
Write-Host ""

# Test which methods actually work
$methods = @(
    @{Name="Credentials"; Test={(Get-Command Connect-PnPOnline).Parameters.ContainsKey("Credentials")}},
    @{Name="Interactive"; Test={(Get-Command Connect-PnPOnline).Parameters.ContainsKey("Interactive")}},
    @{Name="DeviceLogin"; Test={(Get-Command Connect-PnPOnline).Parameters.ContainsKey("DeviceLogin")}},
    @{Name="UseWebLogin"; Test={(Get-Command Connect-PnPOnline).Parameters.ContainsKey("UseWebLogin")}},
    @{Name="LaunchBrowser"; Test={(Get-Command Connect-PnPOnline).Parameters.ContainsKey("LaunchBrowser")}},
    @{Name="ClientId"; Test={(Get-Command Connect-PnPOnline).Parameters.ContainsKey("ClientId")}},
    @{Name="Tenant"; Test={(Get-Command Connect-PnPOnline).Parameters.ContainsKey("Tenant")}}
)

Write-Host "Authentication method availability:" -ForegroundColor Yellow
foreach ($method in $methods) {
    $available = & $method.Test
    $status = if ($available) { "Available" } else { "Not Available" }
    $color = if ($available) { "Green" } else { "Red" }
    Write-Host "  $($method.Name): $status" -ForegroundColor $color
}

Write-Host ""
Write-Host "Recommended authentication approach for your version:" -ForegroundColor Cyan
if ((Get-Command Connect-PnPOnline).Parameters.ContainsKey("Interactive")) {
    Write-Host "  Use -Interactive for browser-based authentication" -ForegroundColor Green
}
elseif ((Get-Command Connect-PnPOnline).Parameters.ContainsKey("UseWebLogin")) {
    Write-Host "  Use -UseWebLogin for browser-based authentication" -ForegroundColor Green
}
else {
    Write-Host "  Use -Credentials with username/password" -ForegroundColor Yellow
    Write-Host "  Note: This may not work with MFA enabled" -ForegroundColor Yellow
}