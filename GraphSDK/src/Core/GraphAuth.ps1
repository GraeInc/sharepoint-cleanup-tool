# GraphAuth.ps1
# Microsoft Graph SDK Authentication Module

function Connect-SharePointGraph {
    <#
    .SYNOPSIS
    Connects to SharePoint using Microsoft Graph SDK
    
    .PARAMETER SiteUrl
    The SharePoint site URL (e.g., https://contoso.sharepoint.com/sites/TeamSite)
    
    .OUTPUTS
    Returns the Graph Site ID if successful, $null otherwise
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl
    )
    
    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        
        # Connect to Graph with required scopes - NO APP REGISTRATION NEEDED!
        Connect-MgGraph -Scopes "Sites.ReadWrite.All", "Sites.Read.All" -NoWelcome -ErrorAction Stop
        
        # Verify connection
        $context = Get-MgContext
        if (-not $context) {
            throw "Failed to establish Graph connection"
        }
        
        Write-Host "Connected as: $($context.Account)" -ForegroundColor Green
        
        # Convert SharePoint URL to Graph Site ID format
        $siteId = ConvertTo-GraphSiteId -Url $SiteUrl
        
        # Verify site access
        $site = Get-MgSite -SiteId $siteId -ErrorAction Stop
        Write-Host "Connected to site: $($site.DisplayName)" -ForegroundColor Green
        
        return @{
            SiteId = $siteId
            Site = $site
            Context = $context
        }
    }
    catch {
        Write-Error "Authentication failed: $_"
        return $null
    }
}

function ConvertTo-GraphSiteId {
    <#
    .SYNOPSIS
    Converts SharePoint URL to Microsoft Graph Site ID format
    
    .PARAMETER Url
    SharePoint site URL
    
    .OUTPUTS
    Graph-compatible site ID string
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url
    )
    
    # Parse URL: https://tenant.sharepoint.com/sites/sitename
    if ($Url -match "https://([^.]+)\.sharepoint\.com(/sites/[^/]+)?") {
        $hostname = "$($matches[1]).sharepoint.com"
        $sitePath = $matches[2]
        
        if ($sitePath) {
            # Site collection: hostname:/sites/sitename
            return "${hostname}:${sitePath}"
        } else {
            # Root site
            return $hostname
        }
    } else {
        throw "Invalid SharePoint URL format"
    }
}

function Disconnect-SharePointGraph {
    <#
    .SYNOPSIS
    Disconnects from Microsoft Graph
    #>
    
    try {
        if (Get-MgContext) {
            Disconnect-MgGraph | Out-Null
            Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Warning "Error during disconnect: $_"
    }
}

function Test-GraphConnection {
    <#
    .SYNOPSIS
    Tests if connected to Microsoft Graph
    
    .OUTPUTS
    Boolean indicating connection status
    #>
    
    try {
        $context = Get-MgContext
        return ($null -ne $context)
    }
    catch {
        return $false
    }
}

# Export functions
Export-ModuleMember -Function Connect-SharePointGraph, ConvertTo-GraphSiteId, Disconnect-SharePointGraph, Test-GraphConnection