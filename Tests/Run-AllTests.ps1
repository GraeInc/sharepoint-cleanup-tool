# Run-AllTests.ps1
# Run all tests for SharePoint Cleanup Tool

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$LibraryName = "Documents",
    
    [Parameter(Mandatory=$true)]
    [datetime]$ModifiedDate = (Get-Date)
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Cleanup Tool - Test Suite" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$testPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$results = @()

# Test 1: Authentication
Write-Host "Running Test 1: Authentication..." -ForegroundColor Yellow
$test1 = & (Join-Path $testPath "Test-Authentication.ps1") -SiteUrl $SiteUrl
$results += @{
    Test = "Authentication"
    Result = if ($LASTEXITCODE -eq 0) { "PASS" } else { "FAIL" }
}
Write-Host ""

# Test 2: Empty Folder Scan
Write-Host "Running Test 2: Empty Folder Scan..." -ForegroundColor Yellow
$test2 = & (Join-Path $testPath "Test-EmptyFolderScan.ps1") -SiteUrl $SiteUrl -LibraryName $LibraryName -ModifiedDate $ModifiedDate
$results += @{
    Test = "Empty Folder Scan"
    Result = if ($LASTEXITCODE -eq 0) { "PASS" } else { "FAIL" }
}
Write-Host ""

# Test 3: CLI Interface
Write-Host "Running Test 3: CLI Interface..." -ForegroundColor Yellow
$test3 = & (Join-Path $testPath "Test-CLI.ps1") -SiteUrl $SiteUrl -LibraryName $LibraryName -ModifiedDate $ModifiedDate
$results += @{
    Test = "CLI Interface"
    Result = if ($LASTEXITCODE -eq 0) { "PASS" } else { "FAIL" }
}
Write-Host ""

# Summary
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "TEST SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$passed = ($results | Where-Object { $_.Result -eq "PASS" }).Count
$failed = ($results | Where-Object { $_.Result -eq "FAIL" }).Count

foreach ($result in $results) {
    $color = if ($result.Result -eq "PASS") { "Green" } else { "Red" }
    Write-Host "$($result.Test): $($result.Result)" -ForegroundColor $color
}

Write-Host ""
Write-Host "Total: $($results.Count) tests" -ForegroundColor White
Write-Host "Passed: $passed" -ForegroundColor Green
Write-Host "Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })

if ($failed -eq 0) {
    Write-Host "`n✓ ALL TESTS PASSED!" -ForegroundColor Green
    exit 0
}
else {
    Write-Host "`n✗ SOME TESTS FAILED" -ForegroundColor Red
    exit 1
}