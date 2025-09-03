# SharePointCleanup-CLI.ps1
# Command-line interface for SharePoint Cleanup Tool

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$LibraryName,
    
    [Parameter(Mandatory=$true)]
    [datetime]$ModifiedDate,
    
    [Parameter()]
    [switch]$WhatIf = $true,
    
    [Parameter()]
    [switch]$ExportResults,
    
    [Parameter()]
    [string]$ExportPath = ".\SharePoint_Empty_Folders_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter()]
    [switch]$Silent
)

# Load core modules
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptPath "src\Core\SharePointManager.ps1")
. (Join-Path $scriptPath "src\Core\Logger.ps1")
. (Join-Path $scriptPath "src\Core\ConfigManager.ps1")

# Initialize modules
$logger = [Logger]::new($scriptPath)
$logger.SetConsoleOutput(-not $Silent)
$config = [ConfigManager]::new($scriptPath)
$spManager = [SharePointManager]::new($SiteUrl)

# Start operation
$logger.LogInfo("===========================================")
$logger.LogInfo("SharePoint Cleanup Tool - CLI Mode")
$logger.LogInfo("===========================================")
$logger.LogInfo("Site: $SiteUrl")
$logger.LogInfo("Library: $LibraryName")
$logger.LogInfo("Date: $($ModifiedDate.ToShortDateString())")
$logger.LogInfo("Mode: $(if ($WhatIf) { 'PREVIEW (No deletions)' } else { 'LIVE (Will delete folders)' })")
$logger.LogInfo("===========================================")

try {
    # Connect to SharePoint
    $logger.LogInfo("Connecting to SharePoint...")
    if (-not $spManager.Connect()) {
        throw "Failed to connect to SharePoint"
    }
    $logger.LogSuccess("Connected successfully")
    $config.SaveRecentSite($SiteUrl)
    
    # Scan for empty folders
    $logger.LogInfo("Scanning for empty folders...")
    $emptyFolders = $spManager.FindEmptyFolders($LibraryName, $ModifiedDate)
    
    $logger.LogInfo("Found $($emptyFolders.Count) empty folders")
    
    if ($emptyFolders.Count -eq 0) {
        $logger.LogInfo("No empty folders found. Exiting.")
        exit 0
    }
    
    # Display results
    if (-not $Silent) {
        Write-Host "`nEmpty Folders Found:" -ForegroundColor Yellow
        Write-Host ("=" * 80)
        $emptyFolders | Format-Table -Property Name, Path, Modified, CreatedBy -AutoSize
        Write-Host ("=" * 80)
    }
    
    # Export if requested
    if ($ExportResults) {
        $logger.LogInfo("Exporting results to: $ExportPath")
        $emptyFolders | Export-Csv -Path $ExportPath -NoTypeInformation
        $logger.LogSuccess("Results exported successfully")
    }
    
    # Delete folders if not in WhatIf mode
    if (-not $WhatIf) {
        $confirmation = if ($Silent) { 
            "Y" 
        } 
        else { 
            Read-Host "`nAre you sure you want to delete $($emptyFolders.Count) folders? This cannot be undone! (Y/N)" 
        }
        
        if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
            $logger.LogInfo("Starting deletion process...")
            $successCount = 0
            $failCount = 0
            
            foreach ($folder in $emptyFolders) {
                try {
                    if ($spManager.DeleteFolder($LibraryName, $folder.Id)) {
                        $logger.LogSuccess("Deleted: $($folder.Name)")
                        $logger.LogDeletion($folder.Name, $folder.Path, $true)
                        $successCount++
                    }
                    else {
                        throw "Delete operation failed"
                    }
                }
                catch {
                    $logger.LogError("Failed to delete: $($folder.Name) - $_")
                    $logger.LogDeletion($folder.Name, $folder.Path, $false)
                    $failCount++
                }
            }
            
            $logger.LogInfo("===========================================")
            $logger.LogInfo("Deletion Summary:")
            $logger.LogSuccess("Successfully deleted: $successCount folders")
            if ($failCount -gt 0) {
                $logger.LogWarning("Failed to delete: $failCount folders")
            }
            $logger.LogInfo("===========================================")
        }
        else {
            $logger.LogInfo("Deletion cancelled by user")
        }
    }
    else {
        $logger.LogInfo("PREVIEW MODE: No folders were deleted")
        $logger.LogInfo("To delete these folders, run without -WhatIf parameter")
    }
    
    $logger.LogInfo("Operation complete. Log saved to: $($logger.GetLogPath())")
}
catch {
    $logger.LogError("Operation failed: $_")
    exit 1
}
finally {
    if ($spManager.IsConnected) {
        $spManager.Disconnect()
    }
}