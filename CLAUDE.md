# SharePoint Cleanup Tool - Development Rules for Claude

## Project Overview
This is a **simple SharePoint cleanup tool** for finding and deleting empty folders. It must be easy for non-technical junior administrators to use.

## Core Requirements
1. **Find empty folders in SharePoint document libraries**
2. **Filter by modified date**
3. **Delete empty folders (with preview mode)**
4. **GUI interface for ease of use**
5. **Audit logging for all operations**

## Technical Stack (UPDATED 2025-09-04)
- **Authentication**: Microsoft Graph PowerShell SDK (NOT PnP.PowerShell)
- **UI**: Windows Forms (System.Windows.Forms)
- **Language**: PowerShell 5.1+ compatible
- **Platform**: Windows 10/11
- **No app registration required** - use Graph SDK's built-in app

## Architecture Rules

### Authentication
- **MUST use Microsoft Graph SDK** (`Connect-MgGraph`)
- **NO PnP.PowerShell** - completely deprecated
- **NO custom app registration** - use SDK's default
- Scopes required: `Sites.ReadWrite.All`
- Support interactive authentication with MFA

### Code Structure
```
SharePointCleaner/
├── SharePointCleaner.ps1       # Main entry (detects GUI/CLI)
├── Install.ps1                 # Module installer
├── src/
│   ├── Core/
│   │   ├── GraphAuth.ps1      # Graph SDK authentication
│   │   ├── FolderOps.ps1      # Folder operations
│   │   └── Logger.ps1         # Logging
│   ├── GUI/
│   │   └── MainGUI.ps1        # Windows Forms GUI
│   └── CLI/
│       └── CLI.ps1            # Command-line interface
├── config/                     # Settings
└── logs/                      # Audit logs
```

### GUI Requirements
- **Three tabs**: Connection, Scan & Clean, Logs
- **Preview mode** enabled by default
- **Progress indicators** for long operations
- **Export to CSV** functionality
- **Recent sites** dropdown
- Must be **responsive** and not freeze

### Logging Requirements
- Log to `logs/sharepoint-cleanup-YYYYMMDD-HHmmss.log`
- Include timestamp, level, and message
- Track all deletions with full path
- Keep audit trail of user actions

## Development Guidelines

### DO's
- ✅ Use Microsoft Graph SDK exclusively
- ✅ Keep the GUI simple and intuitive
- ✅ Test authentication before building features
- ✅ Include progress feedback for users
- ✅ Default to safe mode (preview only)
- ✅ Log everything for audit purposes
- ✅ Support both GUI and CLI modes
- ✅ Handle errors gracefully with user-friendly messages

### DON'Ts
- ❌ Don't use PnP.PowerShell (authentication issues)
- ❌ Don't require app registration
- ❌ Don't make it complex - this is a SIMPLE tool
- ❌ Don't delete without confirmation
- ❌ Don't freeze the GUI during operations
- ❌ Don't assume technical knowledge from users

## Graph SDK Key Commands

### Authentication
```powershell
Connect-MgGraph -Scopes "Sites.ReadWrite.All" -NoWelcome
```

### Site Operations
```powershell
# Get site
$siteId = "contoso.sharepoint.com:/sites/teamsite"
$site = Get-MgSite -SiteId $siteId

# Get lists
$lists = Get-MgSiteList -SiteId $site.Id -Filter "baseTemplate eq 101"

# Get items
$items = Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id

# Delete item
Remove-MgSiteListItem -SiteId $site.Id -ListId $list.Id -ListItemId $item.Id
```

## Testing Checklist
- [ ] Authentication works without app registration
- [ ] Can connect to SharePoint site
- [ ] Can list document libraries
- [ ] Can find folders by date
- [ ] Can identify empty folders
- [ ] Preview mode shows what would be deleted
- [ ] Actual deletion works
- [ ] GUI is responsive
- [ ] Logging captures all operations
- [ ] Works with MFA enabled
- [ ] Works across different tenants

## Error Messages
Keep error messages simple and actionable:
- ❌ "Failed to authenticate: AADSTS700016"
- ✅ "Could not connect to SharePoint. Please check your permissions and try again."

- ❌ "Get-MgSiteListItem : Request failed with status code Unauthorized"
- ✅ "You don't have permission to view this library. Contact your SharePoint administrator."

## Performance Goals
- Authentication: < 10 seconds
- Library scan: < 1 second per 100 folders
- GUI responsiveness: Never freeze
- Memory usage: < 200MB

## User Experience Priority
1. **It must work** - Authentication is critical
2. **It must be simple** - Junior admins are the users
3. **It must be safe** - No accidental deletions
4. **It must log everything** - Audit trail required

## Current Status (2025-09-04 Evening)
- **Migration COMPLETE** from PnP.PowerShell to Microsoft Graph SDK
- **New Issue**: Microsoft Graph SDK v2.30.0 has internal parsing bug
- **Workaround**: Created SharePointCleaner-Direct.ps1 that bypasses the bug
- **Location**: All new code in GraphSDK/ folder

## Known Issues
### Microsoft Graph Module Bug
- **Error**: "The string is missing the terminator: '."
- **Cause**: Internal parsing error in Graph SDK v2.30.0
- **Impact**: Normal module import fails
- **Solution**: Use SharePointCleaner-Direct.ps1 or reinstall older version

## Working Components
- ✅ Authentication logic (when module loads)
- ✅ Site connection
- ✅ Library enumeration  
- ✅ Empty folder detection
- ✅ Folder deletion
- ✅ GUI interface (complete)
- ✅ CLI interface (complete)
- ✅ Logging system

## Next Steps (2025-09-05)
1. Test SharePointCleaner-Direct.ps1 with real SharePoint
2. Try older Graph module versions (2.24.0)
3. If needed, implement REST API fallback
4. Finalize and document working solution
5. Clean up unnecessary files
6. Update main documentation

## Remember
This tool is for **non-technical junior administrators**. If they can't figure it out in 30 seconds, it's too complex. Keep it simple!