try {
    $content = Get-Content 'sharepoint-cleanup-gui.ps1' -Raw
    $null = [scriptblock]::Create($content)
    Write-Host "PowerShell syntax is valid" -ForegroundColor Green
} catch {
    Write-Host "PowerShell syntax error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Error at line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
}
