# Fix-Authentication.ps1
# Fixes authentication issues with PnP.PowerShell v3.x

Write-Host "SharePoint Authentication Fix" -ForegroundColor Cyan
Write-Host "=============================" -ForegroundColor Cyan
Write-Host ""

# Check PnP version
$pnpModule = Get-Module -ListAvailable -Name PnP.PowerShell | Select-Object -First 1
Write-Host "Current PnP.PowerShell version: $($pnpModule.Version)" -ForegroundColor Yellow

# The issue is that PnP v3.x requires explicit authentication configuration
# We'll use the PnP Management Shell App ID which is pre-registered

Write-Host ""
Write-Host "The authentication warning occurs because PnP.PowerShell v3.x requires explicit client ID configuration." -ForegroundColor Yellow
Write-Host ""
Write-Host "SOLUTION OPTIONS:" -ForegroundColor Green
Write-Host ""
Write-Host "Option 1: Use Register-PnPManagementShellAccess (Recommended)" -ForegroundColor Cyan
Write-Host "This registers the PnP Management Shell application for your tenant."
Write-Host ""
Write-Host "Run this command once as admin:" -ForegroundColor Yellow
Write-Host "  Register-PnPManagementShellAccess" -ForegroundColor White
Write-Host ""

Write-Host "Option 2: Use PnP Management Shell Client ID" -ForegroundColor Cyan
Write-Host "Use the built-in PnP Management Shell application ID."
Write-Host ""
Write-Host "Update connection commands to include:" -ForegroundColor Yellow
Write-Host '  -ClientId "31359c7f-bd7e-475c-86db-fdb8c937548e"' -ForegroundColor White
Write-Host ""

Write-Host "Option 3: Use Azure CLI Authentication" -ForegroundColor Cyan
Write-Host "Use Azure CLI for authentication (if installed)."
Write-Host ""
Write-Host "Connect using:" -ForegroundColor Yellow
Write-Host '  Connect-PnPOnline -Url $siteUrl -AzureADWorkloadIdentity' -ForegroundColor White
Write-Host ""

Write-Host "Option 4: Downgrade to PnP.PowerShell 1.12.0" -ForegroundColor Cyan
Write-Host "Use the older version that doesn't require client ID."
Write-Host ""
Write-Host "To downgrade:" -ForegroundColor Yellow
Write-Host "  Uninstall-Module PnP.PowerShell -AllVersions -Force" -ForegroundColor White
Write-Host "  Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Force" -ForegroundColor White
Write-Host ""

# Try to register PnP Management Shell
$register = Read-Host "Would you like to try registering PnP Management Shell access now? (Y/N)"
if ($register -eq 'Y' -or $register -eq 'y') {
    try {
        Write-Host "Attempting to register PnP Management Shell..." -ForegroundColor Yellow
        Register-PnPManagementShellAccess
        Write-Host "Registration successful!" -ForegroundColor Green
        Write-Host "You may need to wait a few minutes for the registration to propagate." -ForegroundColor Yellow
    }
    catch {
        Write-Host "Registration failed: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "You may need to:" -ForegroundColor Yellow
        Write-Host "1. Run PowerShell as Administrator" -ForegroundColor White
        Write-Host "2. Have appropriate Azure AD permissions" -ForegroundColor White
        Write-Host "3. Use a Global Administrator or Application Administrator account" -ForegroundColor White
    }
}