# Test-EmptyFolderScan.ps1
# Test empty folder scanning functionality

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$LibraryName,
    
    [Parameter(Mandatory=$true)]
    [datetime]$ModifiedDate
)

Write-Host "Testing Empty Folder Scan" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host "Site: $SiteUrl" -ForegroundColor White
Write-Host "Library: $LibraryName" -ForegroundColor White
Write-Host "Date: $($ModifiedDate.ToShortDateString())" -ForegroundColor White
Write-Host ""

# Load modules
$scriptPath = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $scriptPath "src\Core\SharePointManager.ps1")
. (Join-Path $scriptPath "src\Core\Logger.ps1")

try {
    # Initialize
    $logger = [Logger]::new($scriptPath)
    $spManager = [SharePointManager]::new($SiteUrl)
    
    # Connect
    Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow
    if (-not $spManager.Connect()) {
        throw "Failed to connect"
    }
    Write-Host "✓ Connected" -ForegroundColor Green
    
    # Get folders by date
    Write-Host "`nGetting folders modified on $($ModifiedDate.ToShortDateString())..." -ForegroundColor Yellow
    $folders = $spManager.GetFoldersByDate($LibraryName, $ModifiedDate)
    Write-Host "✓ Found $($folders.Count) folders" -ForegroundColor Green
    
    # Find empty folders
    Write-Host "`nScanning for empty folders..." -ForegroundColor Yellow
    $emptyFolders = $spManager.FindEmptyFolders($LibraryName, $ModifiedDate)
    Write-Host "✓ Found $($emptyFolders.Count) empty folders" -ForegroundColor Green
    
    # Display results
    if ($emptyFolders.Count -gt 0) {
        Write-Host "`nEmpty folders found:" -ForegroundColor Yellow
        $emptyFolders | Format-Table Name, Path, Modified -AutoSize
    }
    
    # Test logging
    Write-Host "`nTesting logging..." -ForegroundColor Yellow
    foreach ($folder in $emptyFolders | Select-Object -First 3) {
        $logger.LogDeletion($folder.Name, $folder.Path, $true)
    }
    Write-Host "✓ Log file created at: $($logger.GetLogPath())" -ForegroundColor Green
    
    # Disconnect
    $spManager.Disconnect()
    
    Write-Host "`n===================================" -ForegroundColor Cyan
    Write-Host "TEST PASSED" -ForegroundColor Green
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "  Total folders: $($folders.Count)"
    Write-Host "  Empty folders: $($emptyFolders.Count)"
    Write-Host "  Percentage empty: $(if ($folders.Count -gt 0) { [math]::Round(($emptyFolders.Count / $folders.Count) * 100, 2) } else { 0 })%"
    
    exit 0
}
catch {
    Write-Host "✗ TEST FAILED: $_" -ForegroundColor Red
    if ($spManager.IsConnected) {
        $spManager.Disconnect()
    }
    exit 1
}