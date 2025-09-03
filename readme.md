# SharePoint Empty Folder Cleanup Tool

A PowerShell-based tool to identify and remove empty duplicate folders from SharePoint document libraries. Perfect for cleaning up bulk folder creation issues where hundreds or thousands of empty folders were accidentally created.

## 🚀 Quick Start

1. **Download** all files to a folder on your Windows computer
2. **Right-click** on `install-and-run.bat` and select **"Run as administrator"**
3. Follow the prompts to install dependencies and launch the tool

## 📋 Prerequisites

- Windows 10/11 with PowerShell 5.1 or later
- SharePoint Online access with appropriate permissions
- Internet connection for initial setup

## 📁 What's Included

```
📦 sharepoint-cleanup-tool/
├── 📜 install-and-run.bat          # One-click installer and launcher
├── 📄 README.md                    # This documentation
├── 🔧 SharePoint-Cleanup-GUI.ps1   # GUI version of the tool
└── 🔧 SharePoint-Cleanup.ps1       # Command-line version
```

## 🖥️ Tool Options

### Option 1: GUI Tool (Recommended)
- **User-friendly interface** with forms and buttons
- **Visual progress tracking** during scans and deletions  
- **Preview mode** enabled by default for safety
- **Checkbox selection** of folders to delete
- Perfect for one-time cleanup tasks

### Option 2: Command Line Tool
- **Scriptable** for automation
- **Detailed logging** and progress output
- **Flexible parameters** for different scenarios
- **Batch processing** capabilities
- Ideal for recurring maintenance or advanced users

## 📖 How to Use

### Using the GUI Tool

1. **Launch**: Run `install-and-run.bat` and select option 1
2. **Configure**:
   - Enter your SharePoint site URL
   - Specify the document library name
   - Select the date when duplicate folders were created
3. **Scan**: Click "Scan Folders" to find empty folders
4. **Review**: Check/uncheck folders in the results list
5. **Delete**: Uncheck "Preview Mode" and click "Delete Selected"

### Using the Command Line Tool

```powershell
.\SharePoint-Cleanup.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -LibraryName "Documents" -ModifiedDate "2024-01-15"
```

#### Parameters:
- `-SiteUrl`: Your SharePoint site URL
- `-LibraryName`: Name of the document library to clean
- `-ModifiedDate`: Date when the duplicate folders were created
- `-WhatIf`: Set to `$false` to actually delete folders (default: `$true` for safety)

## 🛡️ Safety Features

- **Preview Mode**: Both tools default to safe preview mode
- **Empty Folder Detection**: Only removes truly empty folders
- **Date Filtering**: Only processes folders modified on specific dates
- **Manual Confirmation**: Requires explicit confirmation before deletion
- **Error Handling**: Continues processing even if some folders fail
- **Detailed Logging**: Shows exactly what will be/was deleted

## 🔧 Troubleshooting

### Common Issues and Solutions

| Issue | Solution |
|-------|----------|
| **"PnP.PowerShell module not found"** | Run the installer as administrator |
| **"Access Denied" errors** | Ensure you have SharePoint permissions to delete folders |
| **"Execution policy" warnings** | The installer automatically fixes this |
| **SmartScreen warnings** | Click "More info" then "Run anyway" |
| **No folders found** | Verify the modified date and library name are correct |

### Installation Issues

If the automatic installer fails:

```powershell
# Manual installation steps
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
```

### Authentication Issues

The tool uses interactive authentication. You may see:
- **Browser popup**: Sign in with your SharePoint credentials  
- **Multi-factor authentication**: Complete as required
- **Conditional access**: Follow your organization's policies

## 📊 Example Scenarios

### Scenario 1: Bulk Folder Cleanup
- **Problem**: 1,000 empty folders created on 2024-01-15
- **Solution**: Use GUI tool with date filter
- **Result**: Quick identification and removal of empty folders

### Scenario 2: Regular Maintenance  
- **Problem**: Periodic cleanup needed
- **Solution**: Use command-line tool in scheduled task
- **Result**: Automated maintenance

### Scenario 3: Multiple Libraries
- **Problem**: Several libraries affected
- **Solution**: Run tool multiple times with different library names
- **Result**: Clean multiple locations efficiently

## ⚠️ Important Notes

### Permissions Required
- **SharePoint Site Member** (minimum) with delete permissions
- **Site Owner/Admin** (recommended) for full access

### What Gets Deleted
- ✅ Completely empty folders (no files, no subfolders)
- ✅ Only folders modified on the specified date
- ❌ Folders containing any files or subfolders
- ❌ System folders or special SharePoint folders

### Data Safety
- The tool **cannot recover deleted folders**
- Always test with **Preview Mode** first
- Consider **backing up** important sites before cleanup
- Start with a **small test** on non-critical folders

## 🆘 Support

### Self-Help Resources
1. **Check the date**: Ensure you're using the correct modified date
2. **Verify permissions**: Confirm you can manually delete test folders
3. **Test small batches**: Start with preview mode on a few folders
4. **Check logs**: Review PowerShell output for specific errors

### Getting Additional Help
- Review SharePoint admin center for permissions
- Contact your IT administrator for organizational policies
- Check Microsoft documentation for SharePoint Online limits

## 📝 Technical Details

### System Requirements
- **OS**: Windows 10 1809+ or Windows 11
- **PowerShell**: 5.1 or later (included with Windows)
- **Memory**: Minimal (handles 1000+ folders efficiently)
- **Network**: Internet connection for SharePoint Online

### Dependencies
- **PnP.PowerShell**: Microsoft's official SharePoint PowerShell module
- **Microsoft.PowerShell.Utility**: For GUI components (built-in)
- **System.Windows.Forms**: For GUI interface (built-in)

### Performance
- **Scan Speed**: ~5-10 folders per second
- **Memory Usage**: <100MB for typical workloads
- **Network**: Minimal bandwidth (metadata only)

## 📄 License

This tool is provided as-is for educational and administrative purposes. Use at your own risk and always test in non-production environments first.

---

## 🔄 Version History

- **v1.0**: Initial release with GUI and CLI tools
- Features: Empty folder detection, date filtering, preview mode, batch operations

---

*Need help? Make sure you've read through this README completely and tested with preview mode before making changes to production SharePoint sites.*