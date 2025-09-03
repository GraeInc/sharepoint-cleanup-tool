# SharePoint Empty Folder Cleanup Tool - Integrated Authentication
# Uses current Windows user context or browser session

[CmdletBinding()]
param()

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Cleanup Tool" -ForegroundColor Cyan
Write-Host "Integrated Authentication Version" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if PnP.PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "ERROR: PnP.PowerShell module is not installed." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Import-Module PnP.PowerShell -ErrorAction Stop

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

# Main script
try {
    # Get site URL
    Write-Host "Step 1: Connect to SharePoint" -ForegroundColor Green
    $siteUrl = Get-UserInput -Prompt "Enter SharePoint site URL"
    
    Write-Host ""
    Write-Host "Attempting to connect to SharePoint..." -ForegroundColor Yellow
    Write-Host ""
    
    $connected = $false
    $connectionAttempts = @()
    
    # Method 1: Try current Windows credentials (for on-premises or hybrid)
    if (-not $connected) {
        try {
            Write-Host "Attempting Windows integrated authentication..." -ForegroundColor Yellow
            Connect-PnPOnline -Url $siteUrl -CurrentCredentials -ErrorAction Stop
            Write-Host "Connected using Windows credentials!" -ForegroundColor Green
            $connected = $true
        }
        catch {
            $connectionAttempts += "Windows integrated auth failed: $($_.Exception.Message)"
        }
    }
    
    # Method 2: Try web login with default browser session
    if (-not $connected) {
        try {
            Write-Host "Opening browser for authentication..." -ForegroundColor Yellow
            Write-Host "TIP: If you're already signed in to SharePoint in your browser, this should work automatically." -ForegroundColor Cyan
            Connect-PnPOnline -Url $siteUrl -UseWebLogin -ErrorAction Stop
            Write-Host "Connected via browser!" -ForegroundColor Green
            $connected = $true
        }
        catch {
            $connectionAttempts += "Web login failed: $($_.Exception.Message)"
        }
    }
    
    # Method 3: Try with LaunchBrowser (opens new auth window)
    if (-not $connected) {
        try {
            Write-Host "Attempting browser-based authentication..." -ForegroundColor Yellow
            Connect-PnPOnline -Url $siteUrl -LaunchBrowser -ErrorAction Stop
            Write-Host "Connected!" -ForegroundColor Green
            $connected = $true
        }
        catch {
            $connectionAttempts += "Browser auth failed: $($_.Exception.Message)"
        }
    }
    
    # Method 4: Device login (last resort)
    if (-not $connected) {
        try {
            Write-Host "Attempting device login..." -ForegroundColor Yellow
            Write-Host "You'll need to open a browser and enter a code." -ForegroundColor Cyan
            Connect-PnPOnline -Url $siteUrl -DeviceLogin -ErrorAction Stop
            Write-Host "Connected via device login!" -ForegroundColor Green
            $connected = $true
        }
        catch {
            $connectionAttempts += "Device login failed: $($_.Exception.Message)"
        }
    }
    
    if (-not $connected) {
        Write-Host ""
        Write-Host "ERROR: Could not connect to SharePoint." -ForegroundColor Red
        Write-Host ""
        Write-Host "Connection attempts:" -ForegroundColor Yellow
        foreach ($attempt in $connectionAttempts) {
            Write-Host "  - $attempt" -ForegroundColor Gray
        }
        Write-Host ""
        Write-Host "Suggestions:" -ForegroundColor Yellow
        Write-Host "1. Ensure you have access to the SharePoint site" -ForegroundColor White
        Write-Host "2. Check if the URL is correct" -ForegroundColor White
        Write-Host "3. Try opening the site in your browser first, then run this tool" -ForegroundColor White
        Write-Host "4. Contact your IT administrator for help with authentication" -ForegroundColor White
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    Write-Host ""
    Write-Host "Step 2: Configure Scan Parameters" -ForegroundColor Green
    
    # Get library name
    $libraryName = Get-UserInput -Prompt "Enter document library name" -DefaultValue "Documents"
    
    # Get modified date
    do {
        $dateStr = Get-UserInput -Prompt "Enter modified date (YYYY-MM-DD)"
        try {
            $modifiedDate = [datetime]::ParseExact($dateStr, "yyyy-MM-dd", $null)
            break
        }
        catch {
            Write-Host "Invalid date format. Please use YYYY-MM-DD format." -ForegroundColor Red
        }
    } while ($true)
    
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