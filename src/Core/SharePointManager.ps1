# SharePointManager.ps1
# Core module for SharePoint operations

class SharePointManager {
    [string]$SiteUrl
    [bool]$IsConnected = $false
    [object]$Connection
    
    SharePointManager([string]$url) {
        $this.SiteUrl = $url
    }
    
    # Connect using Windows integrated authentication
    [bool] Connect() {
        try {
            Import-Module PnP.PowerShell -ErrorAction Stop
            
            # Try multiple authentication methods in order of preference
            # For PnP.PowerShell v1.12.0 compatibility
            $tenant = if ($this.SiteUrl -match "https://([^.]+)\.sharepoint\.com") { "$($matches[1]).onmicrosoft.com" } else { $null }
            
            $methods = @(
                @{Name="Interactive"; Script={Connect-PnPOnline -Url $this.SiteUrl -Interactive}},
                @{Name="WebLogin"; Script={Connect-PnPOnline -Url $this.SiteUrl -UseWebLogin}}
            )
            
            foreach ($method in $methods) {
                try {
                    Write-Verbose "Attempting $($method.Name) authentication..."
                    & $method.Script
                    $this.Connection = Get-PnPConnection
                    $this.IsConnected = $true
                    Write-Verbose "Connected successfully using $($method.Name)"
                    return $true
                }
                catch {
                    Write-Verbose "Failed with $($method.Name): $_"
                    continue
                }
            }
            
            throw "All authentication methods failed"
        }
        catch {
            Write-Error "Connection failed: $_"
            return $false
        }
    }
    
    # Get folders modified on specific date
    [array] GetFoldersByDate([string]$libraryName, [datetime]$modifiedDate) {
        if (-not $this.IsConnected) {
            throw "Not connected to SharePoint"
        }
        
        $dateStr = $modifiedDate.ToString('yyyy-MM-dd')
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
                    <Value Type='DateTime'>$dateStr</Value>
                </Eq>
            </And>
        </Where>
    </Query>
</View>
"@
        
        try {
            $folders = Get-PnPListItem -List $libraryName -Query $camlQuery
            return $folders
        }
        catch {
            Write-Error "Failed to get folders: $_"
            return @()
        }
    }
    
    # Check if folder is empty
    [bool] IsFolderEmpty([string]$folderPath) {
        try {
            $contents = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType All -ErrorAction SilentlyContinue
            return ($null -eq $contents -or $contents.Count -eq 0)
        }
        catch {
            Write-Warning "Could not check folder: $folderPath"
            return $false
        }
    }
    
    # Find empty folders
    [array] FindEmptyFolders([string]$libraryName, [datetime]$modifiedDate) {
        $allFolders = $this.GetFoldersByDate($libraryName, $modifiedDate)
        $emptyFolders = @()
        
        foreach ($folder in $allFolders) {
            $folderPath = $folder.FieldValues.FileRef
            $folderName = $folder.FieldValues.FileLeafRef
            
            if ($this.IsFolderEmpty($folderPath)) {
                $emptyFolders += [PSCustomObject]@{
                    Id = $folder.Id
                    Name = $folderName
                    Path = $folderPath
                    Modified = $folder.FieldValues.Modified
                    CreatedBy = $folder.FieldValues.Author.LookupValue
                    ModifiedBy = $folder.FieldValues.Editor.LookupValue
                }
            }
        }
        
        return $emptyFolders
    }
    
    # Delete folder
    [bool] DeleteFolder([string]$libraryName, [int]$folderId) {
        try {
            Remove-PnPListItem -List $libraryName -Identity $folderId -Force
            return $true
        }
        catch {
            Write-Error "Failed to delete folder ID $folderId : $_"
            return $false
        }
    }
    
    # Disconnect
    [void] Disconnect() {
        if ($this.IsConnected) {
            Disconnect-PnPOnline
            $this.IsConnected = $false
        }
    }
}