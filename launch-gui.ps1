# Launcher for SharePoint Cleanup GUI
Write-Host "Starting SharePoint Cleanup GUI..." -ForegroundColor Green

try {
    # Check for the GUI script
    if (Test-Path ".\sharepoint-cleanup-gui-working.ps1") {
        Write-Host "Found GUI script. Launching..." -ForegroundColor Yellow
        & ".\sharepoint-cleanup-gui-working.ps1"
    }
    else {
        Write-Host "ERROR: sharepoint-cleanup-gui-working.ps1 not found!" -ForegroundColor Red
    }
}
catch {
    Write-Host "ERROR launching GUI: $_" -ForegroundColor Red
    Write-Host "Stack trace:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
}

Write-Host ""
Write-Host "GUI closed or failed to launch." -ForegroundColor Yellow
Read-Host "Press Enter to exit"