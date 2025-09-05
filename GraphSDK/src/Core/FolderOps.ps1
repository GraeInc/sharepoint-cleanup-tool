# FolderOps.ps1
# SharePoint Folder Operations using Microsoft Graph SDK

function Get-SharePointLibraries {
    <#
    .SYNOPSIS
    Gets all document libraries from a SharePoint site
    
    .PARAMETER SiteId
    The Graph Site ID
    
    .OUTPUTS
    Array of document library objects
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteId
    )
    
    try {
        # Get all lists, filter for document libraries (BaseTemplate 101)
        $lists = Get-MgSiteList -SiteId $SiteId -Filter "baseTemplate eq 101" -ErrorAction Stop
        
        $libraries = @()
        foreach ($list in $lists) {
            $libraries += @{
                Id = $list.Id
                Name = $list.DisplayName
                WebUrl = $list.WebUrl
                ItemCount = $list.ItemCount
            }
        }
        
        return $libraries
    }
    catch {
        Write-Error "Failed to get libraries: $_"
        return @()
    }
}

function Find-EmptyFolders {
    <#
    .SYNOPSIS
    Finds empty folders in a SharePoint document library
    
    .PARAMETER SiteId
    The Graph Site ID
    
    .PARAMETER LibraryId
    The document library ID
    
    .PARAMETER ModifiedDate
    Filter folders by modified date
    
    .OUTPUTS
    Array of empty folder objects
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteId,
        
        [Parameter(Mandatory=$true)]
        [string]$LibraryId,
        
        [Parameter(Mandatory=$false)]
        [datetime]$ModifiedDate
    )
    
    try {
        Write-Host "Scanning for folders..." -ForegroundColor Yellow
        
        # Get all items from the library
        $allItems = @()
        $pageSize = 200
        
        # Use paging to get all items
        $items = Get-MgSiteListItem -SiteId $SiteId -ListId $LibraryId -Top $pageSize -ExpandProperty "folder,fields"
        $allItems += $items
        
        while ($items.'@odata.nextLink') {
            $items = Invoke-MgGraphRequest -Uri $items.'@odata.nextLink' -Method GET
            $allItems += $items.value
        }
        
        Write-Host "Found $($allItems.Count) total items" -ForegroundColor Cyan
        
        # Separate folders from files
        $folders = $allItems | Where-Object { $_.folder }
        $files = $allItems | Where-Object { -not $_.folder }
        
        Write-Host "Found $($folders.Count) folders to check" -ForegroundColor Cyan
        
        # Build parent-child relationships
        $folderChildren = @{}
        foreach ($item in $allItems) {
            if ($item.parentReference -and $item.parentReference.id) {
                $parentId = $item.parentReference.id
                if (-not $folderChildren.ContainsKey($parentId)) {
                    $folderChildren[$parentId] = @()
                }
                $folderChildren[$parentId] += $item
            }
        }
        
        # Find empty folders
        $emptyFolders = @()
        foreach ($folder in $folders) {
            # Check modified date if specified
            if ($ModifiedDate) {
                $folderDate = [datetime]::Parse($folder.lastModifiedDateTime)
                if ($folderDate.Date -ne $ModifiedDate.Date) {
                    continue
                }
            }
            
            # Check if folder has children
            $hasChildren = $folderChildren.ContainsKey($folder.id)
            
            if (-not $hasChildren) {
                # Extract folder details
                $folderPath = if ($folder.webUrl) { 
                    $folder.webUrl 
                } else { 
                    $folder.name 
                }
                
                $emptyFolders += @{
                    Id = $folder.id
                    Name = $folder.name
                    Path = $folderPath
                    Modified = $folder.lastModifiedDateTime
                    CreatedBy = $folder.createdBy.user.displayName
                    ModifiedBy = $folder.lastModifiedBy.user.displayName
                    Size = 0
                }
            }
        }
        
        Write-Host "Found $($emptyFolders.Count) empty folders" -ForegroundColor Green
        return $emptyFolders
    }
    catch {
        Write-Error "Failed to find empty folders: $_"
        return @()
    }
}

function Remove-EmptyFolder {
    <#
    .SYNOPSIS
    Deletes an empty folder from SharePoint
    
    .PARAMETER SiteId
    The Graph Site ID
    
    .PARAMETER LibraryId
    The document library ID
    
    .PARAMETER FolderId
    The folder item ID to delete
    
    .OUTPUTS
    Boolean indicating success
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteId,
        
        [Parameter(Mandatory=$true)]
        [string]$LibraryId,
        
        [Parameter(Mandatory=$true)]
        [string]$FolderId
    )
    
    try {
        Remove-MgSiteListItem -SiteId $SiteId -ListId $LibraryId -ListItemId $FolderId -ErrorAction Stop
        return $true
    }
    catch {
        Write-Error "Failed to delete folder: $_"
        return $false
    }
}

function Export-FolderReport {
    <#
    .SYNOPSIS
    Exports folder information to CSV
    
    .PARAMETER Folders
    Array of folder objects to export
    
    .PARAMETER Path
    Output file path
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$Folders,
        
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        $exportData = @()
        foreach ($folder in $Folders) {
            $exportData += [PSCustomObject]@{
                Name = $folder.Name
                Path = $folder.Path
                Modified = $folder.Modified
                CreatedBy = $folder.CreatedBy
                ModifiedBy = $folder.ModifiedBy
            }
        }
        
        $exportData | Export-Csv -Path $Path -NoTypeInformation
        Write-Host "Report exported to: $Path" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to export report: $_"
    }
}

# Export functions
Export-ModuleMember -Function Get-SharePointLibraries, Find-EmptyFolders, Remove-EmptyFolder, Export-FolderReport