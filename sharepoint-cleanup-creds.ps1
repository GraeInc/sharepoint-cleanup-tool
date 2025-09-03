# SharePoint Empty Folder Cleanup Tool - Credential Authentication Version
# This version uses username/password authentication (for tenants without app registration)

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$LibraryName,
    
    [Parameter(Mandatory=$false)]
    [string]$ModifiedDate,
    
    [Parameter(Mandatory=$false)]
    [string]$Username,
    
    [switch]$WhatIf = $true
)

# Check if PnP.PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "ERROR: PnP.PowerShell module is not installed." -ForegroundColor Red
    Write-Host "Please run the installer first." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Cleanup Tool (Credential Auth)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Get parameters if not provided
if (-not $SiteUrl) {
    $SiteUrl = Read-Host "Enter SharePoint site URL"
}

if (-not $Username) {
    $Username = Read-Host "Enter your email/username"
}

# Get password securely
$SecurePassword = Read-Host "Enter your password" -AsSecureString

# Create credentials object
$Credentials = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)

Write-Host ""
Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow

try {
    # Try connecting with credentials
    Connect-PnPOnline -Url $SiteUrl -Credentials $Credentials
    Write-Host "Connected successfully!" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "Failed to connect: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Possible issues:" -ForegroundColor Yellow
    Write-Host "1. Multi-factor authentication is enabled (use app registration instead)" -ForegroundColor Yellow
    Write-Host "2. Legacy authentication is disabled in your tenant" -ForegroundColor Yellow
    Write-Host "3. Invalid credentials" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Get library name if not provided
if (-not $LibraryName) {
    $LibraryName = Read-Host "Enter document library name [Documents]"
    if ([string]::IsNullOrWhiteSpace($LibraryName)) {
        $LibraryName = "Documents"
    }
}

# Get date if not provided
if (-not $ModifiedDate) {
    do {
        $dateStr = Read-Host "Enter modified date (YYYY-MM-DD)"
        try {
            $date = [datetime]::ParseExact($dateStr, "yyyy-MM-dd", $null)
            $ModifiedDate = $date.ToString("yyyy-MM-dd")
            break
        }
        catch {
            Write-Host "Invalid date format. Please use YYYY-MM-DD format." -ForegroundColor Red
        }
    } while ($true)
} else {
    $date = [datetime]::Parse($ModifiedDate)
}

Write-Host ""
Write-Host "Scanning for empty folders..." -ForegroundColor Yellow
Write-Host "Library: $LibraryName" -ForegroundColor White
Write-Host "Modified Date: $ModifiedDate" -ForegroundColor White
Write-Host ""

try {
    # Get folders
    $folders = Get-PnPListItem -List $LibraryName -Query "
        <View>
            <Query>
                <Where>
                    <And>
                        <Eq>
                            <FieldRef Name='FSObjType'/>
                            <Value Type='Integer'>1</Value>
                        </Eq>
                        <Eq>
                            <FieldRef Name='Modified'/>
                            <Value Type='DateTime'>$ModifiedDate</Value>
                        </Eq>
                    </And>
                </Where>
            </Query>
        </View>"
    
    Write-Host "Found $($folders.Count) folders modified on $ModifiedDate" -ForegroundColor Cyan
    
    if ($folders.Count -eq 0) {
        Write-Host "No folders found matching the criteria." -ForegroundColor Yellow
        Disconnect-PnPOnline
        Read-Host "Press Enter to exit"
        exit 0
    }
    
    # Check for empty folders
    $emptyFolders = @()
    $processedCount = 0
    
    foreach ($folder in $folders) {
        $processedCount++
        Write-Progress -Activity "Checking folders" -Status "Processing folder $processedCount of $($folders.Count)" -PercentComplete (($processedCount / $folders.Count) * 100)
        
        $folderPath = $folder.FieldValues.FileRef
        $folderName = $folder.FieldValues.FileLeafRef
        
        try {
            $folderContents = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType All -ErrorAction SilentlyContinue
            
            if ($null -eq $folderContents -or $folderContents.Count -eq 0) {
                $emptyFolders += [PSCustomObject]@{
                    Name = $folderName
                    Path = $folderPath
                    Modified = $folder.FieldValues.Modified
                    Id = $folder.Id
                }
            }
        }
        catch {
            # Skip folders we can't access
        }
    }
    
    Write-Progress -Activity "Checking folders" -Completed
    
    # Display results
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "SCAN RESULTS" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Total folders checked: $($folders.Count)" -ForegroundColor White
    Write-Host "Empty folders found: $($emptyFolders.Count)" -ForegroundColor Yellow
    
    if ($emptyFolders.Count -gt 0) {
        Write-Host ""
        Write-Host "Empty folders:" -ForegroundColor Red
        $emptyFolders | Format-Table Name, Modified, Path -AutoSize
        
        if ($WhatIf) {
            Write-Host ""
            Write-Host "*** PREVIEW MODE: No folders were deleted ***" -ForegroundColor Yellow
            Write-Host "To delete, run with -WhatIf:`$false parameter" -ForegroundColor Yellow
        }
        else {
            Write-Host ""
            $confirmation = Read-Host "Delete these $($emptyFolders.Count) empty folders? (Y/N)"
            
            if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
                $deletedCount = 0
                $failedCount = 0
                
                foreach ($emptyFolder in $emptyFolders) {
                    try {
                        Write-Host "Deleting: $($emptyFolder.Name)" -ForegroundColor Gray
                        Remove-PnPListItem -List $LibraryName -Identity $emptyFolder.Id -Force
                        $deletedCount++
                    }
                    catch {
                        Write-Warning "Failed to delete: $($emptyFolder.Name)"
                        $failedCount++
                    }
                }
                
                Write-Host ""
                Write-Host "Deletion complete!" -ForegroundColor Green
                Write-Host "Deleted: $deletedCount folders" -ForegroundColor Green
                if ($failedCount -gt 0) {
                    Write-Host "Failed: $failedCount folders" -ForegroundColor Red
                }
            }
        }
    }
    else {
        Write-Host "No empty folders found." -ForegroundColor Green
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
}

# Disconnect
Write-Host ""
Write-Host "Disconnecting..." -ForegroundColor Yellow
Disconnect-PnPOnline
Write-Host "Done." -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit"