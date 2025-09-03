# SharePoint Cleanup Tool - Simple Runner
# This wrapper uses the working authentication from the CLI script

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Empty Folder Cleanup Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Get parameters interactively
$siteUrl = Read-Host "Enter SharePoint site URL"
$libraryName = Read-Host "Enter library name [Documents]"
if ([string]::IsNullOrWhiteSpace($libraryName)) {
    $libraryName = "Documents"
}

# Get date
do {
    $dateStr = Read-Host "Enter modified date (YYYY-MM-DD)"
    try {
        $modifiedDate = [datetime]::ParseExact($dateStr, "yyyy-MM-dd", $null)
        break
    }
    catch {
        Write-Host "Invalid date format. Please use YYYY-MM-DD" -ForegroundColor Red
    }
} while ($true)

# Preview mode
$preview = Read-Host "Run in preview mode? (Y/N) [Y]"
if ([string]::IsNullOrWhiteSpace($preview)) { $preview = "Y" }
$whatIf = ($preview -eq 'Y' -or $preview -eq 'y')

Write-Host ""
Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  Site: $siteUrl" -ForegroundColor White
Write-Host "  Library: $libraryName" -ForegroundColor White
Write-Host "  Date: $($modifiedDate.ToString('yyyy-MM-dd'))" -ForegroundColor White
Write-Host "  Preview: $whatIf" -ForegroundColor White
Write-Host ""

# Ask about authentication
$useCredentials = Read-Host "Use username/password? (Y/N) [N]"
if ([string]::IsNullOrWhiteSpace($useCredentials)) { $useCredentials = "N" }

Write-Host ""
Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow

# Connect using the EXACT method that works in your CLI script
if ($useCredentials -eq "Y" -or $useCredentials -eq "y") {
    $username = Read-Host "Enter your email/username"
    $SecurePassword = Read-Host "Enter your password" -AsSecureString
    $Credentials = New-Object System.Management.Automation.PSCredential($username, $SecurePassword)
    
    try {
        Connect-PnPOnline -Url $siteUrl -Credentials $Credentials
        Write-Host "Connected successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Credential authentication failed: $_" -ForegroundColor Red
        Write-Host "Trying web login..." -ForegroundColor Yellow
        Connect-PnPOnline -Url $siteUrl -UseWebLogin
    }
}
else {
    Write-Host "A browser window will open for authentication." -ForegroundColor Yellow
    Connect-PnPOnline -Url $siteUrl -UseWebLogin
}

Write-Host ""
Write-Host "Running the cleanup script..." -ForegroundColor Green
Write-Host ""

# Call the original working script with parameters
if ($whatIf) {
    & ".\sharepoint-cleanup-script.ps1" -SiteUrl $siteUrl -LibraryName $libraryName -ModifiedDate $modifiedDate -WhatIf:$true
}
else {
    & ".\sharepoint-cleanup-script.ps1" -SiteUrl $siteUrl -LibraryName $libraryName -ModifiedDate $modifiedDate -WhatIf:$false
}

Write-Host ""
Read-Host "Press Enter to exit"