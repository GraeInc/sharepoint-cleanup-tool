# Register PnP Management Shell App for your tenant
# This will register the app and enable authentication

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PnP App Registration for Your Tenant" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "This script will register the PnP Management Shell app in your tenant." -ForegroundColor Yellow
Write-Host "You need to be a Global Administrator to complete this process." -ForegroundColor Yellow
Write-Host ""

$choice = Read-Host "Do you want to continue? (Y/N)"
if ($choice -ne 'Y' -and $choice -ne 'y') {
    Write-Host "Registration cancelled." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

Write-Host ""
Write-Host "Attempting to register PnP Management Shell..." -ForegroundColor Yellow

try {
    # This will open a browser and register the app
    Register-PnPManagementShellAccess
    
    Write-Host ""
    Write-Host "SUCCESS! The PnP Management Shell has been registered." -ForegroundColor Green
    Write-Host ""
    Write-Host "You can now use the SharePoint Cleanup Tool." -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "Registration failed: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Alternative: You need to manually consent to the app." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Option 1: Admin Consent URL" -ForegroundColor Cyan
    Write-Host "1. Open this URL in your browser (replace YOURTENANT):" -ForegroundColor Yellow
    Write-Host "   https://login.microsoftonline.com/YOURTENANT.onmicrosoft.com/adminconsent?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e" -ForegroundColor White
    Write-Host "2. Sign in as a Global Administrator" -ForegroundColor Yellow
    Write-Host "3. Accept the permissions" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Option 2: Use Credentials Authentication (less secure)" -ForegroundColor Cyan
    Write-Host "The script can be modified to use username/password instead." -ForegroundColor Yellow
    Write-Host ""
}

Read-Host "Press Enter to exit"