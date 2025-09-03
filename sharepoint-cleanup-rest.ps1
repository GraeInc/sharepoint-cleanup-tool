# SharePoint Empty Folder Cleanup Tool - REST API Version
# Uses direct REST API calls with web authentication

[CmdletBinding()]
param()

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Cleanup Tool" -ForegroundColor Cyan
Write-Host "REST API Version (No modules required)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This version uses your default browser's authentication." -ForegroundColor Yellow
Write-Host "Make sure you're signed into SharePoint in Internet Explorer or Edge." -ForegroundColor Yellow
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

# Function to make authenticated REST calls
function Invoke-SPRestMethod {
    param(
        [string]$Url,
        [string]$Method = "GET",
        [object]$Body = $null
    )
    
    try {
        $params = @{
            Uri = $Url
            Method = $Method
            UseDefaultCredentials = $true
            ContentType = "application/json;odata=verbose"
            Headers = @{
                "Accept" = "application/json;odata=verbose"
            }
        }
        
        if ($Body) {
            $params.Body = $Body | ConvertTo-Json -Depth 10
        }
        
        # Use Internet Explorer's session
        $response = Invoke-WebRequest @params -UseBasicParsing
        return $response.Content | ConvertFrom-Json
    }
    catch {
        # Try with Invoke-RestMethod as fallback
        try {
            $response = Invoke-RestMethod @params
            return $response
        }
        catch {
            throw $_
        }
    }
}

try {
    # Get site URL
    Write-Host "Step 1: Connect to SharePoint" -ForegroundColor Green
    $siteUrl = Get-UserInput -Prompt "Enter SharePoint site URL"
    
    # Test connection
    Write-Host ""
    Write-Host "Testing connection to SharePoint..." -ForegroundColor Yellow
    
    try {
        $webUrl = "$siteUrl/_api/web"
        $web = Invoke-SPRestMethod -Url $webUrl
        Write-Host "Connected to: $($web.d.Title)" -ForegroundColor Green
        Write-Host ""
    }
    catch {
        Write-Host "Failed to connect to SharePoint." -ForegroundColor Red
        Write-Host ""
        Write-Host "Please try the following:" -ForegroundColor Yellow
        Write-Host "1. Open Internet Explorer or Edge" -ForegroundColor White
        Write-Host "2. Navigate to: $siteUrl" -ForegroundColor White
        Write-Host "3. Sign in to SharePoint" -ForegroundColor White
        Write-Host "4. Keep the browser open and run this script again" -ForegroundColor White
        Write-Host ""
        Write-Host "Error: $_" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Get parameters
    Write-Host "Step 2: Configure Scan Parameters" -ForegroundColor Green
    $libraryName = Get-UserInput -Prompt "Enter document library name" -DefaultValue "Documents"
    
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
    
    # Get folders
    Write-Host ""
    Write-Host "Step 3: Scanning for Empty Folders" -ForegroundColor Green
    Write-Host "Getting folders from library: $libraryName" -ForegroundColor Yellow
    
    # Build CAML query
    $dateFormatted = $modifiedDate.ToString("yyyy-MM-dd")
    $camlQuery = @"
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
                    <Value Type='DateTime'>$dateFormatted</Value>
                </Eq>
            </And>
        </Where>
    </Query>
    <RowLimit>5000</RowLimit>
</View>
"@
    
    # Get list items using CAML
    $listUrl = "$siteUrl/_api/web/lists/getbytitle('$libraryName')/getitems"
    $body = @{
        query = @{
            ViewXml = $camlQuery
        }
    }
    
    try {
        $response = Invoke-SPRestMethod -Url $listUrl -Method "POST" -Body $body
        $folders = $response.d.results
        
        Write-Host "Found $($folders.Count) folders modified on $($modifiedDate.ToShortDateString())" -ForegroundColor Cyan
        
        if ($folders.Count -eq 0) {
            Write-Host "No folders found matching the criteria." -ForegroundColor Yellow
            Read-Host "Press Enter to exit"
            exit 0
        }
    }
    catch {
        Write-Host "Error getting folders: $_" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Check for empty folders
    Write-Host "Checking folders for content..." -ForegroundColor Yellow
    $emptyFolders = @()
    $processedCount = 0
    
    foreach ($folder in $folders) {
        $processedCount++
        Write-Progress -Activity "Checking folders" -Status "Processing folder $processedCount of $($folders.Count)" -PercentComplete (($processedCount / $folders.Count) * 100)
        
        try {
            # Get folder contents
            $folderUrl = "$siteUrl/_api/web/GetFolderByServerRelativeUrl('$($folder.FileRef)')/Files"
            $files = Invoke-SPRestMethod -Url $folderUrl
            
            $subfolderUrl = "$siteUrl/_api/web/GetFolderByServerRelativeUrl('$($folder.FileRef)')/Folders"
            $subfolders = Invoke-SPRestMethod -Url $subfolderUrl
            
            # Filter out system folders
            $realSubfolders = $subfolders.d.results | Where-Object { $_.Name -ne "Forms" }
            
            if ($files.d.results.Count -eq 0 -and $realSubfolders.Count -eq 0) {
                $emptyFolders += [PSCustomObject]@{
                    Name = $folder.FileLeafRef
                    Path = $folder.FileRef
                    Modified = $folder.Modified
                    Id = $folder.ID
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
            $confirmation = Get-UserInput -Prompt "Delete these $($emptyFolders.Count) empty folders? (Y/N)" -DefaultValue "N"
            
            if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
                Write-Host ""
                Write-Host "Deleting empty folders..." -ForegroundColor Red
                $deletedCount = 0
                $failedCount = 0
                
                # Get request digest for delete operations
                $contextUrl = "$siteUrl/_api/contextinfo"
                $contextInfo = Invoke-SPRestMethod -Url $contextUrl -Method "POST"
                $digest = $contextInfo.d.GetContextWebInformation.FormDigestValue
                
                foreach ($emptyFolder in $emptyFolders) {
                    try {
                        Write-Host "Deleting: $($emptyFolder.Name)" -ForegroundColor Gray
                        
                        $deleteUrl = "$siteUrl/_api/web/lists/getbytitle('$libraryName')/items($($emptyFolder.Id))"
                        
                        Invoke-WebRequest -Uri $deleteUrl -Method DELETE `
                            -UseDefaultCredentials `
                            -Headers @{
                                "Accept" = "application/json;odata=verbose"
                                "X-RequestDigest" = $digest
                                "IF-MATCH" = "*"
                                "X-HTTP-Method" = "DELETE"
                            }
                        
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
            else {
                Write-Host "Deletion cancelled." -ForegroundColor Yellow
            }
        }
    }
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit"