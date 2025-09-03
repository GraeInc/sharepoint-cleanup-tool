# Test-CLI.ps1
# Test CLI interface

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$LibraryName,
    
    [Parameter(Mandatory=$true)]
    [datetime]$ModifiedDate
)

Write-Host "Testing CLI Interface" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host ""

$scriptPath = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$cliScript = Join-Path $scriptPath "SharePointCleanup-CLI.ps1"

if (-not (Test-Path $cliScript)) {
    Write-Host "✗ CLI script not found at: $cliScript" -ForegroundColor Red
    exit 1
}

try {
    Write-Host "Test 1: Preview mode (WhatIf)" -ForegroundColor Yellow
    & $cliScript -SiteUrl $SiteUrl -LibraryName $LibraryName -ModifiedDate $ModifiedDate -WhatIf -Silent
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Preview mode test passed" -ForegroundColor Green
    }
    else {
        throw "Preview mode test failed"
    }
    
    Write-Host "`nTest 2: Export results" -ForegroundColor Yellow
    $exportFile = ".\test_export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    & $cliScript -SiteUrl $SiteUrl -LibraryName $LibraryName -ModifiedDate $ModifiedDate -WhatIf -ExportResults -ExportPath $exportFile -Silent
    
    if (Test-Path $exportFile) {
        Write-Host "✓ Export test passed - File created: $exportFile" -ForegroundColor Green
        Remove-Item $exportFile -Force
    }
    else {
        throw "Export test failed - File not created"
    }
    
    Write-Host "`n===================================" -ForegroundColor Cyan
    Write-Host "CLI TESTS PASSED" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "✗ TEST FAILED: $_" -ForegroundColor Red
    exit 1
}