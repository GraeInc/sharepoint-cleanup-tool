# SharePoint Empty Folder Cleanup Script
# Requires PnP.PowerShell module: Install-Module -Name PnP.PowerShell

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$LibraryName,
    
    [Parameter(Mandatory=$true)]
    [datetime]$ModifiedDate,
    
    [switch]$WhatIf = $true  # Safe mode by default - set to $false to actually delete
)

# Connect to SharePoint
Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Green

# Check if we should use credential authentication
$useCredentials = Read-Host "Use username/password authentication? (Y/N) [Y]"
if ([string]::IsNullOrWhiteSpace($useCredentials)) { $useCredentials = "Y" }

if ($useCredentials -eq "Y" -or $useCredentials -eq "y") {
    # Get credentials
    $username = Read-Host "Enter your email/username"
    $SecurePassword = Read-Host "Enter your password" -AsSecureString
    $Credentials = New-Object System.Management.Automation.PSCredential($username, $SecurePassword)
    
    try {
        Connect-PnPOnline -Url $SiteUrl -Credentials $Credentials
        Write-Host "Connected successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Credential authentication failed: $_" -ForegroundColor Red
        Write-Host "Trying web login..." -ForegroundColor Yellow
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    }
}
else {
    Write-Host "A browser window will open for authentication." -ForegroundColor Yellow
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
}

try {
    # Get all folders from the document library
    Write-Host "Getting folders from library: $LibraryName" -ForegroundColor Yellow
    
    # Get folders modified on the specific date
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
                            <Value Type='DateTime'>$($ModifiedDate.ToString('yyyy-MM-dd'))</Value>
                        </Eq>
                    </And>
                </Where>
            </Query>
        </View>"
    
    Write-Host "Found $($folders.Count) folders modified on $($ModifiedDate.ToShortDateString())" -ForegroundColor Cyan
    
    $emptyFolders = @()
    $processedCount = 0
    
    foreach ($folder in $folders) {
        $processedCount++
        Write-Progress -Activity "Checking folders" -Status "Processing folder $processedCount of $($folders.Count)" -PercentComplete (($processedCount / $folders.Count) * 100)
        
        $folderPath = $folder.FieldValues.FileRef
        $folderName = $folder.FieldValues.FileLeafRef
        
        Write-Host "Checking folder: $folderName" -ForegroundColor Gray
        
        try {
            # Check if folder is empty (no files or subfolders)
            $folderContents = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType All -ErrorAction SilentlyContinue
            
            if ($folderContents.Count -eq 0) {
                $emptyFolders += [PSCustomObject]@{
                    Name = $folderName
                    Path = $folderPath
                    Modified = $folder.FieldValues.Modified
                    Id = $folder.Id
                }
                Write-Host "  → Empty folder found: $folderName" -ForegroundColor Red
            }
            else {
                Write-Host "  → Folder contains $($folderContents.Count) items" -ForegroundColor Green
            }
        }
        catch {
            Write-Warning "Could not access folder: $folderName - $($_.Exception.Message)"
        }
    }
    
    Write-Progress -Activity "Checking folders" -Completed
    
    Write-Host "`n=== SUMMARY ===" -ForegroundColor Magenta
    Write-Host "Total folders checked: $($folders.Count)" -ForegroundColor White
    Write-Host "Empty folders found: $($emptyFolders.Count)" -ForegroundColor Yellow
    
    if ($emptyFolders.Count -gt 0) {
        Write-Host "`nEmpty folders to be removed:" -ForegroundColor Red
        $emptyFolders | Format-Table Name, Modified, Path -AutoSize
        
        if ($WhatIf) {
            Write-Host "`n*** WHAT-IF MODE: No folders will be deleted ***" -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "To actually delete these folders, run the script with -WhatIf:`$false" -ForegroundColor Yellow
        }
        else {
            $confirmation = Read-Host "`nDo you want to delete these $($emptyFolders.Count) empty folders? (y/N)"
            
            if ($confirmation -eq 'y' -or $confirmation -eq 'Y') {
                $deletedCount = 0
                
                foreach ($emptyFolder in $emptyFolders) {
                    try {
                        Write-Host "Deleting: $($emptyFolder.Name)" -ForegroundColor Red
                        Remove-PnPListItem -List $LibraryName -Identity $emptyFolder.Id -Force
                        $deletedCount++
                    }
                    catch {
                        Write-Error "Failed to delete $($emptyFolder.Name): $($_.Exception.Message)"
                    }
                }
                
                Write-Host "`nSuccessfully deleted $deletedCount out of $($emptyFolders.Count) empty folders." -ForegroundColor Green
            }
            else {
                Write-Host "Operation cancelled." -ForegroundColor Yellow
            }
        }
    }
    else {
        Write-Host "No empty folders found with the specified criteria." -ForegroundColor Green
    }
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
finally {
    Disconnect-PnPOnline
}

# Example usage:
# .\SharePoint-Cleanup.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -LibraryName "Documents" -ModifiedDate "2024-01-15" -WhatIf:$false