# CLI.ps1
# Command-line interface for SharePoint Cleanup Tool

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$LibraryName,
    
    [Parameter(Mandatory=$true)]
    [datetime]$ModifiedDate,
    
    [switch]$WhatIf = $true,
    
    [switch]$ExportResults,
    
    [string]$ExportPath = ".\EmptyFolders_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [switch]$Silent
)

# Import core modules
$corePath = Join-Path $PSScriptRoot "..\Core"
. (Join-Path $corePath "GraphAuth.ps1")
. (Join-Path $corePath "FolderOps.ps1")
. (Join-Path $corePath "Logger.ps1")

function Run-CLICleanup {
    try {
        if (-not $Silent) {
            Write-Host ""
            Write-Host "SharePoint Cleanup Tool - Graph SDK CLI" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "Site URL: $SiteUrl" -ForegroundColor Yellow
            Write-Host "Library: $LibraryName" -ForegroundColor Yellow
            Write-Host "Date Filter: $($ModifiedDate.ToString('yyyy-MM-dd'))" -ForegroundColor Yellow
            Write-Host "Mode: $(if ($WhatIf) { 'PREVIEW' } else { 'LIVE' })" -ForegroundColor $(if ($WhatIf) { 'Green' } else { 'Red' })
            Write-Host ""
        }
        
        Write-Log "INFO" "CLI mode started"
        Write-Log "INFO" "Parameters: Site=$SiteUrl, Library=$LibraryName, Date=$($ModifiedDate.ToString('yyyy-MM-dd')), WhatIf=$WhatIf"
        
        # Connect to SharePoint
        if (-not $Silent) { Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow }
        $connection = Connect-SharePointGraph -SiteUrl $SiteUrl
        
        if (-not $connection) {
            Write-Log "ERROR" "Failed to connect to SharePoint"
            if (-not $Silent) { Write-Host "Failed to connect to SharePoint" -ForegroundColor Red }
            exit 1
        }
        
        # Get library
        if (-not $Silent) { Write-Host "Getting document library..." -ForegroundColor Yellow }
        $libraries = Get-SharePointLibraries -SiteId $connection.SiteId
        $library = $libraries | Where-Object { $_.Name -eq $LibraryName } | Select-Object -First 1
        
        if (-not $library) {
            Write-Log "ERROR" "Library not found: $LibraryName"
            if (-not $Silent) { Write-Host "Library not found: $LibraryName" -ForegroundColor Red }
            Disconnect-SharePointGraph
            exit 1
        }
        
        # Find empty folders
        if (-not $Silent) { Write-Host "Scanning for empty folders..." -ForegroundColor Yellow }
        $emptyFolders = Find-EmptyFolders -SiteId $connection.SiteId -LibraryId $library.Id -ModifiedDate $ModifiedDate
        
        if ($emptyFolders.Count -eq 0) {
            Write-Log "INFO" "No empty folders found"
            if (-not $Silent) { 
                Write-Host ""
                Write-Host "No empty folders found matching criteria" -ForegroundColor Green 
            }
            Disconnect-SharePointGraph
            exit 0
        }
        
        # Display results
        if (-not $Silent) {
            Write-Host ""
            Write-Host "Found $($emptyFolders.Count) empty folder(s):" -ForegroundColor Green
            Write-Host ""
            
            $emptyFolders | ForEach-Object {
                Write-Host "  - $($_.Name)" -ForegroundColor Yellow
                Write-Host "    Path: $($_.Path)" -ForegroundColor Gray
                Write-Host "    Modified: $($_.Modified) by $($_.ModifiedBy)" -ForegroundColor Gray
            }
            Write-Host ""
        }
        
        # Export if requested
        if ($ExportResults) {
            Export-FolderReport -Folders $emptyFolders -Path $ExportPath
            Write-Log "INFO" "Results exported to: $ExportPath"
        }
        
        # Delete folders
        if ($WhatIf) {
            Write-Log "INFO" "Preview mode - no folders will be deleted"
            if (-not $Silent) {
                Write-Host "PREVIEW MODE: No folders will be deleted" -ForegroundColor Green
                Write-Host "Remove -WhatIf parameter to perform actual deletion" -ForegroundColor Yellow
            }
        } else {
            $confirmMessage = "Are you sure you want to delete $($emptyFolders.Count) folder(s)? This cannot be undone! (Y/N)"
            $confirm = if ($Silent) { "Y" } else { Read-Host $confirmMessage }
            
            if ($confirm -eq "Y" -or $confirm -eq "y") {
                $deleted = 0
                $failed = 0
                
                foreach ($folder in $emptyFolders) {
                    if (Remove-EmptyFolder -SiteId $connection.SiteId -LibraryId $library.Id -FolderId $folder.Id) {
                        Write-Log "DELETE-SUCCESS" "Deleted: $($folder.Name)"
                        if (-not $Silent) { Write-Host "  Deleted: $($folder.Name)" -ForegroundColor Green }
                        $deleted++
                    } else {
                        Write-Log "DELETE-FAIL" "Failed to delete: $($folder.Name)"
                        if (-not $Silent) { Write-Host "  Failed: $($folder.Name)" -ForegroundColor Red }
                        $failed++
                    }
                }
                
                Write-Log "INFO" "Deletion complete: Success=$deleted, Failed=$failed"
                if (-not $Silent) {
                    Write-Host ""
                    Write-Host "Deletion complete: $deleted succeeded, $failed failed" -ForegroundColor Cyan
                }
            } else {
                Write-Log "INFO" "User cancelled deletion"
                if (-not $Silent) { Write-Host "Deletion cancelled" -ForegroundColor Yellow }
            }
        }
        
        # Disconnect
        Disconnect-SharePointGraph
        
        if (-not $Silent) {
            Write-Host ""
            Write-Host "Operation complete" -ForegroundColor Green
            Write-Host "Log file: $(Get-LogPath)" -ForegroundColor Gray
        }
        
        exit 0
    }
    catch {
        Write-Log "ERROR" "Unexpected error: $_"
        if (-not $Silent) { 
            Write-Host ""
            Write-Host "Error: $_" -ForegroundColor Red 
        }
        
        if (Test-GraphConnection) {
            Disconnect-SharePointGraph
        }
        
        exit 1
    }
}

# Run CLI
Run-CLICleanup