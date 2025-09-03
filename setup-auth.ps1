# PnP PowerShell Authentication Setup Script
# This script helps set up authentication for PnP PowerShell v3.x

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PnP PowerShell Authentication Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "This script will help you set up authentication for SharePoint access." -ForegroundColor Yellow
Write-Host ""

# Check if PnP.PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell module is not installed." -ForegroundColor Red
    Write-Host "Please run the installer first." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Import-Module PnP.PowerShell

Write-Host "Choose an authentication method:" -ForegroundColor Green
Write-Host "1. Use PnP Management Shell (Recommended - no setup required)" -ForegroundColor White
Write-Host "2. Register PnP Management Shell for your tenant (one-time setup)" -ForegroundColor White
Write-Host "3. Instructions for creating custom Azure AD app" -ForegroundColor White
Write-Host ""

$choice = Read-Host "Enter your choice (1-3)"

switch ($choice) {
    "1" {
        Write-Host ""
        Write-Host "The scripts are already configured to use the PnP Management Shell app." -ForegroundColor Green
        Write-Host "Client ID: 31359c7f-bd7e-475c-86db-fdb8c937548e" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "This should work for most users without any additional setup." -ForegroundColor Yellow
        Write-Host "Just run the cleanup tool and authenticate when prompted." -ForegroundColor Yellow
    }
    
    "2" {
        Write-Host ""
        Write-Host "Registering PnP Management Shell for your tenant..." -ForegroundColor Yellow
        Write-Host "You will need to authenticate as a Global Administrator." -ForegroundColor Yellow
        Write-Host ""
        
        try {
            Register-PnPManagementShellAccess
            Write-Host ""
            Write-Host "Registration successful!" -ForegroundColor Green
            Write-Host "The PnP Management Shell is now registered for your tenant." -ForegroundColor Green
        }
        catch {
            Write-Host "Registration failed: $_" -ForegroundColor Red
            Write-Host ""
            Write-Host "You may need to:" -ForegroundColor Yellow
            Write-Host "1. Run this script as an administrator" -ForegroundColor Yellow
            Write-Host "2. Ensure you have Global Administrator permissions" -ForegroundColor Yellow
        }
    }
    
    "3" {
        Write-Host ""
        Write-Host "To create a custom Azure AD app registration:" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "1. Go to https://portal.azure.com" -ForegroundColor White
        Write-Host "2. Navigate to Azure Active Directory > App registrations" -ForegroundColor White
        Write-Host "3. Click 'New registration'" -ForegroundColor White
        Write-Host "4. Name: SharePoint Cleanup Tool" -ForegroundColor White
        Write-Host "5. Supported account types: Single tenant" -ForegroundColor White
        Write-Host "6. Redirect URI: http://localhost" -ForegroundColor White
        Write-Host "7. After creation, go to 'API permissions'" -ForegroundColor White
        Write-Host "8. Add permissions:" -ForegroundColor White
        Write-Host "   - SharePoint > Delegated > AllSites.FullControl" -ForegroundColor Cyan
        Write-Host "   - SharePoint > Delegated > User.Read.All" -ForegroundColor Cyan
        Write-Host "9. Grant admin consent for your organization" -ForegroundColor White
        Write-Host "10. Copy the Application (client) ID" -ForegroundColor White
        Write-Host ""
        Write-Host "Then update the scripts to use your Client ID:" -ForegroundColor Yellow
        Write-Host 'Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId "YOUR-CLIENT-ID-HERE"' -ForegroundColor Cyan
    }
    
    default {
        Write-Host "Invalid choice." -ForegroundColor Red
    }
}

Write-Host ""
Read-Host "Press Enter to exit"