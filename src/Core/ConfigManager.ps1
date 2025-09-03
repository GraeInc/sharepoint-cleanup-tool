# ConfigManager.ps1
# Configuration management module

class ConfigManager {
    [hashtable]$Config
    [string]$ConfigPath
    
    ConfigManager([string]$basePath) {
        $configDir = Join-Path $basePath "Config"
        if (-not (Test-Path $configDir)) {
            New-Item -ItemType Directory -Path $configDir -Force | Out-Null
        }
        
        $this.ConfigPath = Join-Path $configDir "settings.json"
        $this.LoadConfig()
    }
    
    [void] LoadConfig() {
        if (Test-Path $this.ConfigPath) {
            try {
                $jsonContent = Get-Content $this.ConfigPath -Raw
                $this.Config = $jsonContent | ConvertFrom-Json -AsHashtable
            }
            catch {
                $this.Config = $this.GetDefaultConfig()
                $this.SaveConfig()
            }
        }
        else {
            $this.Config = $this.GetDefaultConfig()
            $this.SaveConfig()
        }
    }
    
    [hashtable] GetDefaultConfig() {
        return @{
            DefaultLibrary = "Documents"
            PreviewMode = $true
            MaxBatchSize = 100
            EnableLogging = $true
            LogRetentionDays = 30
            LastUsedSite = ""
            AuthenticationMethod = "Auto"
            ShowConfirmations = $true
        }
    }
    
    [void] SaveConfig() {
        $this.Config | ConvertTo-Json -Depth 10 | Set-Content $this.ConfigPath
    }
    
    [object] Get([string]$key) {
        return $this.Config[$key]
    }
    
    [void] Set([string]$key, [object]$value) {
        $this.Config[$key] = $value
        $this.SaveConfig()
    }
    
    [hashtable] GetAll() {
        return $this.Config
    }
    
    [void] SaveRecentSite([string]$siteUrl) {
        if (-not $this.Config.ContainsKey('RecentSites')) {
            $this.Config['RecentSites'] = @()
        }
        
        # Remove if already exists
        $this.Config['RecentSites'] = @($this.Config['RecentSites'] | Where-Object { $_ -ne $siteUrl })
        
        # Add to beginning
        $this.Config['RecentSites'] = @($siteUrl) + $this.Config['RecentSites']
        
        # Keep only last 5
        if ($this.Config['RecentSites'].Count -gt 5) {
            $this.Config['RecentSites'] = $this.Config['RecentSites'][0..4]
        }
        
        $this.Config['LastUsedSite'] = $siteUrl
        $this.SaveConfig()
    }
    
    [array] GetRecentSites() {
        if ($this.Config.ContainsKey('RecentSites')) {
            return $this.Config['RecentSites']
        }
        return @()
    }
}