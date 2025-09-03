# SharePoint Empty Folder Cleanup Simple GUI Tool
# Requires PnP.PowerShell module
# Version: 3.0 - Simplified for compatibility
# Author: SharePoint Cleanup Tool

[CmdletBinding()]
param()

# Check if PnP.PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "ERROR: PnP.PowerShell module is not installed." -ForegroundColor Red
    Write-Host "Please run: Install-Module -Name PnP.PowerShell" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Import module
Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Empty Folder Cleanup Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Function to get user input
function Get-UserInput {
    param(
        [string]$Prompt,
        [string]$DefaultValue = ""
    )
    
    if ($DefaultValue) {
        $input = Read-Host "$Prompt [$DefaultValue]"
        if ([string]::IsNullOrWhiteSpace($input)) {
            return $DefaultValue
        }
        return $input
    }
    else {
        do {
            $input = Read-Host $Prompt
        } while ([string]::IsNullOrWhiteSpace($input))
        return $input
    }
}

# Function to get date input
function Get-DateInput {
    param(
        [string]$Prompt
    )
    
    do {
        $dateStr = Read-Host "$Prompt (YYYY-MM-DD)"
        try {
            $date = [datetime]::ParseExact($dateStr, "yyyy-MM-dd", $null)
            return $date
        }
        catch {
            Write-Host "Invalid date format. Please use YYYY-MM-DD format." -ForegroundColor Red
        }
    } while ($true)
}

# Main script
try {
    # Get site URL
    Write-Host "Step 1: Connect to SharePoint" -ForegroundColor Green
    $siteUrl = Get-UserInput -Prompt "Enter SharePoint site URL"
    
    # Get credentials
    $username = Get-UserInput -Prompt "Enter your email/username"
    $SecurePassword = Read-Host "Enter your password" -AsSecureString
    $Credentials = New-Object System.Management.Automation.PSCredential($username, $SecurePassword)
    
    # Connect to SharePoint
    Write-Host ""
    Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow
    Write-Host ""
    
    try {
        # Try credential authentication first
        Connect-PnPOnline -Url $siteUrl -Credentials $Credentials
        Write-Host "Connected successfully!" -ForegroundColor Green
        Write-Host ""
    }
    catch {
        Write-Host "Credential authentication failed: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "Trying interactive authentication..." -ForegroundColor Yellow
        
        try {
            # Fallback to interactive without app ID (may work for some tenants)
            Connect-PnPOnline -Url $siteUrl -UseWebLogin
            Write-Host "Connected successfully!" -ForegroundColor Green
            Write-Host ""
        }
        catch {
            Write-Host "Failed to connect to SharePoint: $_" -ForegroundColor Red
            Write-Host ""
            Write-Host "Possible issues:" -ForegroundColor Yellow
            Write-Host "1. Multi-factor authentication may be blocking credential auth" -ForegroundColor Yellow
            Write-Host "2. Legacy authentication might be disabled" -ForegroundColor Yellow
            Write-Host "3. Invalid credentials or URL" -ForegroundColor Yellow
            Write-Host "4. Try running: Register-PnPManagementShellAccess" -ForegroundColor Yellow
            Read-Host "Press Enter to exit"
            exit 1
        }
    }
    
    # Get library name
    Write-Host "Step 2: Configure Scan Parameters" -ForegroundColor Green
    $libraryName = Get-UserInput -Prompt "Enter document library name" -DefaultValue "Documents"
    
    # Get modified date
    $modifiedDate = Get-DateInput -Prompt "Enter modified date to search for"
    
    # Ask for preview mode
    Write-Host ""
    $previewResponse = Get-UserInput -Prompt "Run in preview mode? (Y/N)" -DefaultValue "Y"
    $previewMode = $previewResponse -eq 'Y' -or $previewResponse -eq 'y'
    
    if ($previewMode) {
        Write-Host "Preview mode: ON - No folders will be deleted" -ForegroundColor Yellow
    }
    else {
        Write-Host "Preview mode: OFF - Folders will be deleted after confirmation" -ForegroundColor Red
    }
    
    # Scan for empty folders
    Write-Host ""
    Write-Host "Step 3: Scanning for Empty Folders" -ForegroundColor Green
    Write-Host "Getting folders from library: $libraryName" -ForegroundColor Yellow
    
    try {
        $folders = Get-PnPListItem -List $libraryName -Query "
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
                                <Value Type='DateTime'>$($modifiedDate.ToString('yyyy-MM-dd'))</Value>
                            </Eq>
                        </And>
                    </Where>
                </Query>
            </View>"
        
        Write-Host "Found $($folders.Count) folders modified on $($modifiedDate.ToShortDateString())" -ForegroundColor Cyan
        
        if ($folders.Count -eq 0) {
            Write-Host "No folders found matching the criteria." -ForegroundColor Yellow
            Disconnect-PnPOnline
            Read-Host "Press Enter to exit"
            exit 0
        }
        
        # Check for empty folders
        Write-Host "Checking folders for content..." -ForegroundColor Yellow
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
        
        if ($emptyFolders.Count -eq 0) {
            Write-Host ""
            Write-Host "No empty folders found." -ForegroundColor Green
        }
        else {
            Write-Host ""
            Write-Host "Empty folders:" -ForegroundColor Red
            $emptyFolders | Format-Table Name, Modified, Path -AutoSize
            
            if ($previewMode) {
                Write-Host ""
                Write-Host "*** PREVIEW MODE: No folders were deleted ***" -ForegroundColor Yellow
                Write-Host "To delete these folders, run again with preview mode OFF" -ForegroundColor Yellow
            }
            else {
                Write-Host ""
                $confirmation = Get-UserInput -Prompt "Do you want to delete these $($emptyFolders.Count) empty folders? (Y/N)" -DefaultValue "N"
                
                if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
                    Write-Host ""
                    Write-Host "Deleting empty folders..." -ForegroundColor Red
                    $deletedCount = 0
                    $failedCount = 0
                    
                    foreach ($emptyFolder in $emptyFolders) {
                        try {
                            Write-Host "Deleting: $($emptyFolder.Name)" -ForegroundColor Gray
                            Remove-PnPListItem -List $libraryName -Identity $emptyFolder.Id -Force
                            $deletedCount++
                        }
                        catch {
                            Write-Warning "Failed to delete: $($emptyFolder.Name) - $_"
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
                else {
                    Write-Host "Deletion cancelled." -ForegroundColor Yellow
                }
            }
        }
    }
    catch {
        Write-Host "Error scanning folders: $_" -ForegroundColor Red
    }
    
    # Disconnect
    Write-Host ""
    Write-Host "Disconnecting from SharePoint..." -ForegroundColor Yellow
    Disconnect-PnPOnline
    Write-Host "Disconnected." -ForegroundColor Green
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit"