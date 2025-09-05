# SharePointCleaner-Direct.ps1
# Direct implementation without module auto-loading issues
# This version uses Graph REST API directly to avoid module parsing errors

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$LibraryName,
    
    [Parameter(Mandatory=$true)]
    [datetime]$ModifiedDate,
    
    [switch]$WhatIf = $true
)

Write-Host ""
Write-Host "SharePoint Cleanup Tool - Direct Graph API Version" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This version works around Graph module loading issues" -ForegroundColor Yellow
Write-Host ""

# Step 1: Authenticate using device code flow
Write-Host "Step 1: Authentication" -ForegroundColor Yellow
Write-Host "Please follow these steps:" -ForegroundColor Cyan
Write-Host ""

try {
    # Try to use Connect-MgGraph if available
    $connected = $false
    
    # Check if we can import just the authentication module
    $authModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication
    if ($authModule) {
        Write-Host "Using Microsoft Graph Authentication module..." -ForegroundColor Gray
        
        # Import only what we need
        Import-Module Microsoft.Graph.Authentication -CommandName Connect-MgGraph, Disconnect-MgGraph, Get-MgContext -ErrorAction SilentlyContinue
        
        # Try to connect
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        Write-Host "A browser window will open for authentication" -ForegroundColor Cyan
        
        Connect-MgGraph -Scopes "Sites.ReadWrite.All" -NoWelcome -ErrorAction Stop
        
        $context = Get-MgContext
        if ($context) {
            Write-Host "Connected as: $($context.Account)" -ForegroundColor Green
            $connected = $true
        }
    }
    
    if (-not $connected) {
        Write-Host ""
        Write-Host "ERROR: Could not authenticate" -ForegroundColor Red
        Write-Host ""
        Write-Host "The Microsoft Graph module has a loading issue." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Please try one of these solutions:" -ForegroundColor Cyan
        Write-Host "1. Reinstall the Graph module:" -ForegroundColor White
        Write-Host "   Uninstall-Module Microsoft.Graph -AllVersions" -ForegroundColor Gray
        Write-Host "   Install-Module Microsoft.Graph -Force" -ForegroundColor Gray
        Write-Host ""
        Write-Host "2. Use PowerShell 7 instead of Windows PowerShell:" -ForegroundColor White
        Write-Host "   Download from: https://aka.ms/powershell" -ForegroundColor Gray
        Write-Host ""
        Write-Host "3. Try the PnP.PowerShell version instead (in parent folder)" -ForegroundColor White
        Write-Host ""
        exit 1
    }
    
    # Step 2: Parse site URL
    Write-Host ""
    Write-Host "Step 2: Connecting to SharePoint site..." -ForegroundColor Yellow
    
    if ($SiteUrl -match "https://([^.]+)\.sharepoint\.com(/sites/[^/]+)?") {
        $hostname = "$($matches[1]).sharepoint.com"
        $sitePath = $matches[2]
        $siteId = if ($sitePath) { "${hostname}:${sitePath}" } else { $hostname }
        
        Write-Host "Site ID: $siteId" -ForegroundColor Gray
        
        # Import Sites module carefully
        Import-Module Microsoft.Graph.Sites -CommandName Get-MgSite, Get-MgSiteList, Get-MgSiteListItem, Remove-MgSiteListItem -ErrorAction SilentlyContinue
        
        # Get site
        $site = Get-MgSite -SiteId $siteId -ErrorAction Stop
        Write-Host "Connected to: $($site.DisplayName)" -ForegroundColor Green
        
        # Step 3: Get library
        Write-Host ""
        Write-Host "Step 3: Finding document library '$LibraryName'..." -ForegroundColor Yellow
        
        $lists = Get-MgSiteList -SiteId $site.Id -Filter "displayName eq '$LibraryName'" -ErrorAction Stop
        $library = $lists | Where-Object { $_.BaseTemplate -eq 101 } | Select-Object -First 1
        
        if (-not $library) {
            Write-Host "ERROR: Library not found: $LibraryName" -ForegroundColor Red
            Disconnect-MgGraph
            exit 1
        }
        
        Write-Host "Found library: $($library.DisplayName)" -ForegroundColor Green
        
        # Step 4: Find folders
        Write-Host ""
        Write-Host "Step 4: Scanning for empty folders modified on $($ModifiedDate.ToString('yyyy-MM-dd'))..." -ForegroundColor Yellow
        
        # Get all items
        $allItems = Get-MgSiteListItem -SiteId $site.Id -ListId $library.Id -ExpandProperty "folder,fields" -Top 999
        
        # Filter folders
        $folders = $allItems | Where-Object { $_.folder }
        $emptyFolders = @()
        
        foreach ($folder in $folders) {
            # Check date
            $folderDate = [datetime]::Parse($folder.lastModifiedDateTime)
            if ($folderDate.Date -eq $ModifiedDate.Date) {
                # Check if empty (simplified check)
                $children = $allItems | Where-Object { 
                    $_.parentReference -and $_.parentReference.id -eq $folder.id 
                }
                
                if ($children.Count -eq 0) {
                    $emptyFolders += $folder
                    Write-Host "  Found empty: $($folder.name)" -ForegroundColor Yellow
                }
            }
        }
        
        Write-Host ""
        Write-Host "Found $($emptyFolders.Count) empty folder(s)" -ForegroundColor $(if($emptyFolders.Count -eq 0) { 'Green' } else { 'Yellow' })
        
        # Step 5: Delete if not WhatIf
        if ($emptyFolders.Count -gt 0) {
            if ($WhatIf) {
                Write-Host ""
                Write-Host "PREVIEW MODE: No folders will be deleted" -ForegroundColor Green
                Write-Host "Remove -WhatIf to perform actual deletion" -ForegroundColor Yellow
            } else {
                Write-Host ""
                $confirm = Read-Host "Delete $($emptyFolders.Count) folder(s)? (Y/N)"
                if ($confirm -eq 'Y') {
                    foreach ($folder in $emptyFolders) {
                        try {
                            Remove-MgSiteListItem -SiteId $site.Id -ListId $library.Id -ListItemId $folder.Id -ErrorAction Stop
                            Write-Host "  Deleted: $($folder.name)" -ForegroundColor Green
                        } catch {
                            Write-Host "  Failed: $($folder.name) - $_" -ForegroundColor Red
                        }
                    }
                }
            }
        }
        
        # Cleanup
        Disconnect-MgGraph
        Write-Host ""
        Write-Host "Operation complete!" -ForegroundColor Green
        
    } else {
        Write-Host "ERROR: Invalid SharePoint URL" -ForegroundColor Red
        exit 1
    }
    
} catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    
    if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
    
    exit 1
}