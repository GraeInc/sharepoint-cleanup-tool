# SharePoint Empty Folder Cleanup Tool - Microsoft Graph Version
# Uses Microsoft Graph PowerShell SDK which has better authentication support

[CmdletBinding()]
param()

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Cleanup Tool" -ForegroundColor Cyan
Write-Host "Microsoft Graph Version" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if Microsoft.Graph is installed
$graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Sites
if (-not $graphModule) {
    Write-Host "Microsoft Graph PowerShell module is not installed." -ForegroundColor Yellow
    Write-Host ""
    $install = Read-Host "Would you like to install it now? (Y/N)"
    if ($install -eq 'Y' -or $install -eq 'y') {
        Write-Host "Installing Microsoft Graph PowerShell modules..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Sites -Scope CurrentUser -Force
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
        Write-Host "Installation complete!" -ForegroundColor Green
    }
    else {
        Write-Host "Module is required to continue. Exiting." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Import-Module Microsoft.Graph.Sites -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

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

try {
    # Connect to Microsoft Graph
    Write-Host "Step 1: Connect to Microsoft Graph" -ForegroundColor Green
    Write-Host "A browser window will open for authentication." -ForegroundColor Yellow
    Write-Host ""
    
    # Connect with minimal permissions needed
    Connect-MgGraph -Scopes "Sites.Read.All", "Sites.ReadWrite.All" -NoWelcome
    
    $context = Get-MgContext
    if ($context) {
        Write-Host "Connected successfully as: $($context.Account)" -ForegroundColor Green
        Write-Host ""
    }
    else {
        throw "Failed to connect to Microsoft Graph"
    }
    
    # Get site URL
    Write-Host "Step 2: Specify SharePoint Site" -ForegroundColor Green
    $siteUrl = Get-UserInput -Prompt "Enter SharePoint site URL"
    
    # Parse the site URL to get site ID
    Write-Host "Getting site information..." -ForegroundColor Yellow
    
    # Extract hostname and site path from URL
    $uri = [System.Uri]$siteUrl
    $hostname = $uri.Host
    $sitePath = $uri.AbsolutePath
    
    # Get the site
    try {
        if ($sitePath -match "^/sites/(.+)$") {
            $siteName = $matches[1]
            $site = Get-MgSite -Search $siteName | Where-Object { $_.WebUrl -eq $siteUrl } | Select-Object -First 1
        }
        else {
            # Try root site
            $site = Get-MgSite -SiteId "$hostname:/"
        }
        
        if (-not $site) {
            throw "Site not found"
        }
        
        Write-Host "Found site: $($site.DisplayName)" -ForegroundColor Green
        Write-Host ""
    }
    catch {
        Write-Host "Error: Could not find site. $_" -ForegroundColor Red
        Write-Host "Make sure you have access to the site and the URL is correct." -ForegroundColor Yellow
        Disconnect-MgGraph
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Get document library
    Write-Host "Step 3: Configure Scan Parameters" -ForegroundColor Green
    $libraryName = Get-UserInput -Prompt "Enter document library name" -DefaultValue "Documents"
    
    # Get the library (drive)
    Write-Host "Getting document library..." -ForegroundColor Yellow
    $drives = Get-MgSiteDrive -SiteId $site.Id
    $library = $drives | Where-Object { $_.Name -like "*$libraryName*" } | Select-Object -First 1
    
    if (-not $library) {
        Write-Host "Library '$libraryName' not found. Available libraries:" -ForegroundColor Yellow
        $drives | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor White }
        Disconnect-MgGraph
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    Write-Host "Found library: $($library.Name)" -ForegroundColor Green
    
    # Get date
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
    
    # Preview mode
    Write-Host ""
    $previewResponse = Get-UserInput -Prompt "Run in preview mode? (Y/N)" -DefaultValue "Y"
    $previewMode = $previewResponse -eq 'Y' -or $previewResponse -eq 'y'
    
    if ($previewMode) {
        Write-Host "Preview mode: ON - No folders will be deleted" -ForegroundColor Yellow
    }
    else {
        Write-Host "Preview mode: OFF - Folders will be deleted after confirmation" -ForegroundColor Red
    }
    
    # Get all items in the library
    Write-Host ""
    Write-Host "Step 4: Scanning for Empty Folders" -ForegroundColor Green
    Write-Host "Getting folder list..." -ForegroundColor Yellow
    
    # Get root items
    $allItems = @()
    $rootItems = Get-MgDriveItem -DriveId $library.Id
    $allItems += Get-MgDriveItemChild -DriveId $library.Id -DriveItemId $rootItems.Id -All
    
    # Filter for folders
    $folders = $allItems | Where-Object { 
        $_.Folder -ne $null -and 
        $_.LastModifiedDateTime.Date -eq $modifiedDate.Date 
    }
    
    Write-Host "Found $($folders.Count) folders modified on $($modifiedDate.ToShortDateString())" -ForegroundColor Cyan
    
    if ($folders.Count -eq 0) {
        Write-Host "No folders found matching the criteria." -ForegroundColor Yellow
        Disconnect-MgGraph
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
        
        try {
            # Get children of the folder
            $children = Get-MgDriveItemChild -DriveId $library.Id -DriveItemId $folder.Id
            
            if ($children.Count -eq 0) {
                $emptyFolders += [PSCustomObject]@{
                    Name = $folder.Name
                    Path = $folder.WebUrl
                    Modified = $folder.LastModifiedDateTime
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
        $emptyFolders | Format-Table Name, Modified -AutoSize
        
        if ($previewMode) {
            Write-Host ""
            Write-Host "*** PREVIEW MODE: No folders were deleted ***" -ForegroundColor Yellow
            Write-Host "To delete these folders, run again with preview mode OFF" -ForegroundColor Yellow
        }
        else {
            Write-Host ""
            $confirmation = Get-UserInput -Prompt "Delete these $($emptyFolders.Count) empty folders? (Y/N)" -DefaultValue "N"
            
            if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
                Write-Host ""
                Write-Host "Deleting empty folders..." -ForegroundColor Red
                $deletedCount = 0
                $failedCount = 0
                
                foreach ($emptyFolder in $emptyFolders) {
                    try {
                        Write-Host "Deleting: $($emptyFolder.Name)" -ForegroundColor Gray
                        Remove-MgDriveItem -DriveId $library.Id -DriveItemId $emptyFolder.Id
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
    
    # Disconnect
    Write-Host ""
    Write-Host "Disconnecting..." -ForegroundColor Yellow
    Disconnect-MgGraph
    Write-Host "Done." -ForegroundColor Green
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    if (Get-MgContext) {
        Disconnect-MgGraph
    }
}

Write-Host ""
Read-Host "Press Enter to exit"