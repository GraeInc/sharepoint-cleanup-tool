# SharePoint Cleanup Tool v2.0

A comprehensive enterprise tool for managing empty folders in SharePoint Online with full audit logging, Windows authentication, and both GUI and CLI interfaces.

## 🎯 Key Features

- **Windows Integrated Authentication**: Supports multiple auth methods (DeviceLogin, Interactive, WebLogin)
- **Dual Interface**: Full-featured GUI and scriptable CLI
- **Audit Logging**: Complete activity logging with timestamps and outcomes
- **Safe by Default**: Preview mode enabled by default
- **Batch Operations**: Process hundreds of folders efficiently
- **Export Capabilities**: Export results to CSV for reporting
- **Configuration Management**: Saves settings and recent sites

## 📦 Installation

### Quick Start
1. Run `Install.bat` as Administrator
2. Double-click `Run-GUI.bat` to launch the GUI

### Manual Installation
```powershell
# Install PnP.PowerShell module
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force

# Set execution policy if needed
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 🖥️ GUI Mode

### Launching the GUI
```powershell
.\SharePointCleanup.ps1
# OR
.\Run-GUI.bat
```

### GUI Features
- **Tabbed Interface**:
  - Connection tab for authentication
  - Scan & Clean tab for operations
  - Activity Log tab for audit trail
- **Recent Sites**: Quick access to previously used sites
- **Visual Progress**: Progress bars and status indicators
- **Data Grid**: Select/deselect folders for deletion
- **Export Results**: Save scan results to CSV

### GUI Workflow
1. **Connect**: Enter SharePoint site URL and click Connect
2. **Authenticate**: Complete browser/device authentication
3. **Configure**: Select library and date filter
4. **Scan**: Click "Scan Folders" to find empty folders
5. **Review**: Check results in the data grid
6. **Delete**: Uncheck "Preview Mode" and click "Delete Selected"

## 📝 CLI Mode

### Basic Usage
```powershell
.\SharePointCleanup.ps1 -CLI -SiteUrl <url> -LibraryName <name> -ModifiedDate <date>
```

### Examples
```powershell
# Preview mode (default)
.\SharePointCleanup.ps1 -CLI -SiteUrl "https://contoso.sharepoint.com/sites/mysite" `
    -LibraryName "Documents" -ModifiedDate "2024-01-15"

# Live mode (actual deletion)
.\SharePointCleanup.ps1 -CLI -SiteUrl "https://contoso.sharepoint.com/sites/mysite" `
    -LibraryName "Documents" -ModifiedDate "2024-01-15" -WhatIf:$false

# Export results
.\SharePointCleanup-CLI.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/mysite" `
    -LibraryName "Documents" -ModifiedDate "2024-01-15" `
    -ExportResults -ExportPath ".\results.csv"

# Silent mode for automation
.\SharePointCleanup-CLI.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/mysite" `
    -LibraryName "Documents" -ModifiedDate "2024-01-15" `
    -Silent -WhatIf:$false
```

### CLI Parameters
- `-SiteUrl`: SharePoint site URL (required)
- `-LibraryName`: Document library name (required)
- `-ModifiedDate`: Date filter for folders (required)
- `-WhatIf`: Preview mode flag (default: $true)
- `-ExportResults`: Export results to CSV
- `-ExportPath`: Path for export file
- `-Silent`: Suppress console output

## 📊 Logging

### Log Location
Logs are automatically saved to:
```
.\Logs\sharepoint-cleanup-YYYYMMDD-HHmmss.log
```

### Log Contents
- Connection events
- Authentication attempts
- Scan operations
- Deletion attempts (success/failure)
- Error messages
- User actions

### Log Format
```
[2025-01-15 10:30:45] [INFO] Connected to SharePoint: https://contoso.sharepoint.com
[2025-01-15 10:31:02] [SUCCESS] Scan complete - Found 125 empty folders
[2025-01-15 10:31:45] [DELETE-SUCCESS] Folder: DuplicateFolder001 | Path: /sites/mysite/Documents/DuplicateFolder001
```

## 🧪 Testing

### Run All Tests
```powershell
.\Tests\Run-AllTests.ps1 -SiteUrl <url> -LibraryName <library> -ModifiedDate <date>
```

### Individual Tests
```powershell
# Test authentication
.\Tests\Test-Authentication.ps1 -SiteUrl <url>

# Test scanning
.\Tests\Test-EmptyFolderScan.ps1 -SiteUrl <url> -LibraryName <library> -ModifiedDate <date>

# Test CLI
.\Tests\Test-CLI.ps1 -SiteUrl <url> -LibraryName <library> -ModifiedDate <date>
```

## 🔒 Security

### Authentication Methods
1. **Device Login**: Provides a code to enter on microsoft.com/devicelogin
2. **Interactive**: Opens browser for direct login
3. **Web Login**: Legacy browser authentication

### Permissions Required
- Minimum: SharePoint site member with delete permissions
- Recommended: Site owner or admin

### Safety Features
- Preview mode by default
- Confirmation prompts before deletion
- Only deletes completely empty folders
- Full audit logging
- No recursive deletion

## 📁 Project Structure

```
SharepointCleaner/
├── Install.bat                 # Installer script
├── Run-GUI.bat                # GUI launcher
├── SharePointCleanup.ps1      # Main entry point
├── SharePointCleanup-CLI.ps1  # CLI interface
├── src/
│   ├── Core/
│   │   ├── SharePointManager.ps1  # SharePoint operations
│   │   ├── Logger.ps1            # Logging module
│   │   └── ConfigManager.ps1     # Configuration management
│   └── GUI/
│       └── MainForm.ps1          # GUI application
├── Tests/
│   ├── Run-AllTests.ps1         # Test suite runner
│   ├── Test-Authentication.ps1   # Auth tests
│   ├── Test-EmptyFolderScan.ps1 # Scan tests
│   └── Test-CLI.ps1             # CLI tests
├── Config/                       # Configuration files (auto-created)
└── Logs/                        # Log files (auto-created)
```

## ⚙️ Configuration

### Settings File
Configuration is stored in `.\Config\settings.json`:
```json
{
  "DefaultLibrary": "Documents",
  "PreviewMode": true,
  "MaxBatchSize": 100,
  "EnableLogging": true,
  "LogRetentionDays": 30,
  "LastUsedSite": "https://contoso.sharepoint.com/sites/mysite",
  "RecentSites": []
}
```

### Modifying Settings
Settings are automatically saved when using the GUI. For CLI, modify the JSON file directly.

## 🚀 Advanced Usage

### Automation Script
```powershell
# Daily cleanup script
$sites = @(
    "https://contoso.sharepoint.com/sites/site1",
    "https://contoso.sharepoint.com/sites/site2"
)

foreach ($site in $sites) {
    .\SharePointCleanup-CLI.ps1 -SiteUrl $site `
        -LibraryName "Documents" `
        -ModifiedDate (Get-Date).AddDays(-1) `
        -WhatIf:$false -Silent
}
```

### Scheduled Task
```powershell
# Create scheduled task for daily cleanup
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-File C:\Tools\SharePointCleanup\SharePointCleanup-CLI.ps1 -SiteUrl https://contoso.sharepoint.com -LibraryName Documents -ModifiedDate (Get-Date) -WhatIf:`$false -Silent"

$trigger = New-ScheduledTaskTrigger -Daily -At 2am

Register-ScheduledTask -TaskName "SharePointCleanup" `
    -Action $action -Trigger $trigger `
    -Description "Daily SharePoint empty folder cleanup"
```

## 🐛 Troubleshooting

### Common Issues

| Issue | Solution |
|-------|----------|
| Authentication fails | Try different auth method (DeviceLogin usually works best) |
| Module not found | Run Install.bat as Administrator |
| Access denied | Check SharePoint permissions |
| No folders found | Verify date format (YYYY-MM-DD) |
| GUI won't launch | Check PowerShell execution policy |

### Debug Mode
```powershell
# Enable verbose output
$VerbosePreference = "Continue"
.\SharePointCleanup.ps1 -CLI -SiteUrl <url> -LibraryName <lib> -ModifiedDate <date>
```

## 📈 Performance

- Scans ~10-20 folders per second
- Handles 1000+ folders efficiently
- Memory usage: <100MB typical
- Network: Minimal bandwidth (metadata only)

## 📝 License

This tool is provided for administrative use. Ensure compliance with your organization's policies before use.

## 🤝 Support

For issues or feature requests, contact your IT administrator or check the logs for detailed error messages.

---

**Version**: 2.0  
**Last Updated**: January 2025  
**Requirements**: Windows 10/11, PowerShell 5.1+, PnP.PowerShell module