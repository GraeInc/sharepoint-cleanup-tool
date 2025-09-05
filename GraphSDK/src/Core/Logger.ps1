# Logger.ps1
# Audit Logging Module

$Script:LogFile = $null
$Script:LogPath = Join-Path $PSScriptRoot "..\..\logs"

function Initialize-Logger {
    <#
    .SYNOPSIS
    Initializes the logger with a new log file
    #>
    
    # Ensure log directory exists
    if (-not (Test-Path $Script:LogPath)) {
        New-Item -ItemType Directory -Path $Script:LogPath -Force | Out-Null
    }
    
    # Create log file name with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $Script:LogFile = Join-Path $Script:LogPath "sharepoint-cleanup-$timestamp.log"
    
    # Write header
    Write-Log "INFO" "SharePoint Cleanup Tool - Graph SDK Version"
    Write-Log "INFO" "Session started by: $env:USERNAME"
    Write-Log "INFO" "Computer: $env:COMPUTERNAME"
}

function Write-Log {
    <#
    .SYNOPSIS
    Writes a message to the log file and optionally to console
    
    .PARAMETER Level
    Log level (INFO, WARNING, ERROR, SUCCESS, DELETE)
    
    .PARAMETER Message
    Log message
    
    .PARAMETER NoConsole
    Don't output to console
    #>
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DELETE", "DELETE-SUCCESS", "DELETE-FAIL")]
        [string]$Level,
        
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [switch]$NoConsole
    )
    
    # Ensure logger is initialized
    if (-not $Script:LogFile) {
        Initialize-Logger
    }
    
    # Create log entry
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to file
    try {
        Add-Content -Path $Script:LogFile -Value $logEntry -ErrorAction Stop
    }
    catch {
        # If can't write to log, at least output to console
        Write-Host "Failed to write to log: $_" -ForegroundColor Red
    }
    
    # Write to console unless suppressed
    if (-not $NoConsole) {
        switch ($Level) {
            "INFO"           { Write-Host $Message -ForegroundColor Cyan }
            "WARNING"        { Write-Host $Message -ForegroundColor Yellow }
            "ERROR"          { Write-Host $Message -ForegroundColor Red }
            "SUCCESS"        { Write-Host $Message -ForegroundColor Green }
            "DELETE"         { Write-Host $Message -ForegroundColor Magenta }
            "DELETE-SUCCESS" { Write-Host $Message -ForegroundColor Green }
            "DELETE-FAIL"    { Write-Host $Message -ForegroundColor Red }
        }
    }
}

function Get-LogContent {
    <#
    .SYNOPSIS
    Gets the content of the current log file
    
    .OUTPUTS
    String array of log entries
    #>
    
    if ($Script:LogFile -and (Test-Path $Script:LogFile)) {
        return Get-Content -Path $Script:LogFile
    }
    return @()
}

function Get-LogPath {
    <#
    .SYNOPSIS
    Gets the current log file path
    
    .OUTPUTS
    Path to the current log file
    #>
    
    return $Script:LogFile
}

# Initialize on module load
Initialize-Logger

# Export functions
Export-ModuleMember -Function Initialize-Logger, Write-Log, Get-LogContent, Get-LogPath