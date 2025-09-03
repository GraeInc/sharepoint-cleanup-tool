# Logger.ps1
# Logging module for audit trail

class Logger {
    [string]$LogPath
    [string]$SessionId
    [bool]$ConsoleOutput = $true
    
    Logger([string]$basePath) {
        $this.SessionId = (Get-Date -Format "yyyyMMdd-HHmmss")
        $logDir = Join-Path $basePath "Logs"
        
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        
        $this.LogPath = Join-Path $logDir "sharepoint-cleanup-$($this.SessionId).log"
        $this.WriteLog("INFO", "Logging session started")
    }
    
    [void] WriteLog([string]$level, [string]$message) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] [$level] $message"
        
        # Write to file
        Add-Content -Path $this.LogPath -Value $logEntry
        
        # Write to console if enabled
        if ($this.ConsoleOutput) {
            switch ($level) {
                "ERROR" { Write-Host $logEntry -ForegroundColor Red }
                "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
                "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
                "INFO" { Write-Host $logEntry -ForegroundColor Cyan }
                default { Write-Host $logEntry }
            }
        }
    }
    
    [void] LogInfo([string]$message) {
        $this.WriteLog("INFO", $message)
    }
    
    [void] LogSuccess([string]$message) {
        $this.WriteLog("SUCCESS", $message)
    }
    
    [void] LogWarning([string]$message) {
        $this.WriteLog("WARNING", $message)
    }
    
    [void] LogError([string]$message) {
        $this.WriteLog("ERROR", $message)
    }
    
    [void] LogAction([string]$action, [hashtable]$details) {
        $detailsJson = $details | ConvertTo-Json -Compress
        $this.WriteLog("ACTION", "$action | Details: $detailsJson")
    }
    
    [void] LogDeletion([string]$folderName, [string]$folderPath, [bool]$success) {
        $status = if ($success) { "SUCCESS" } else { "FAILED" }
        $this.WriteLog("DELETE-$status", "Folder: $folderName | Path: $folderPath")
    }
    
    [string] GetLogPath() {
        return $this.LogPath
    }
    
    [void] SetConsoleOutput([bool]$enabled) {
        $this.ConsoleOutput = $enabled
    }
}