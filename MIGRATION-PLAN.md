# SharePoint Cleanup Tool - Complete Graph SDK Migration Plan
*Date: 2025-09-04*

## Core Requirements (from README)
1. **Simple tool for SharePoint cleanup** - Find and delete empty folders
2. **GUI for non-technical users** - Junior admins need easy interface  
3. **Windows integrated authentication** - Support MFA, no manual credentials
4. **Date filtering** - Find folders modified on specific date
5. **Audit logging** - Track all operations
6. **Preview mode by default** - Safe operations

## Current Problems with PnP.PowerShell
- Authentication fails with tenant restrictions (client ID not registered)
- Requires workarounds that aren't sustainable
- Not truly multi-tenant without app registration
- Complex authentication flow

## Microsoft Graph SDK Advantages
- ✅ Built-in multi-tenant app (no registration needed)
- ✅ Native MFA support
- ✅ Works across all tenants
- ✅ Actively maintained by Microsoft
- ✅ Single authentication flow

## Implementation Plan

### Phase 1: Core Architecture (Graph SDK)
```
SharePointCleaner-GraphSDK/
├── SharePointCleaner.ps1          # Main launcher (GUI/CLI detection)
├── Install.ps1                    # Installs Graph SDK modules
├── src/
│   ├── Core/
│   │   ├── GraphAuthentication.ps1    # Handle Graph SDK auth
│   │   ├── SharePointOperations.ps1   # Empty folder operations
│   │   └── Logger.ps1                 # Audit logging
│   ├── GUI/
│   │   └── CleanupGUI.ps1            # Windows Forms GUI
│   └── CLI/
│       └── CleanupCLI.ps1            # Command-line interface
├── config/
│   └── settings.json                  # User preferences
└── logs/                             # Audit logs
```

### Phase 2: Core Functions Mapping

| PnP.PowerShell | Microsoft Graph SDK | Purpose |
|----------------|-------------------|----------|
| Connect-PnPOnline | Connect-MgGraph | Authentication |
| Get-PnPWeb | Get-MgSite | Get site info |
| Get-PnPList | Get-MgSiteList | Get document libraries |
| Get-PnPListItem | Get-MgSiteListItem | Get folders |
| Remove-PnPListItem | Remove-MgSiteListItem | Delete folders |
| Disconnect-PnPOnline | Disconnect-MgGraph | Cleanup |

### Phase 3: Key Implementation Details

#### Authentication (GraphAuthentication.ps1)
```powershell
function Connect-SharePointGraph {
    param([string]$SiteUrl)
    
    # Simple authentication - no app registration needed!
    Connect-MgGraph -Scopes "Sites.ReadWrite.All" -NoWelcome
    
    # Parse site URL to Graph format
    $siteId = ConvertTo-GraphSiteId -Url $SiteUrl
    return $siteId
}
```

#### Finding Empty Folders (SharePointOperations.ps1)
```powershell
function Find-EmptyFolders {
    param(
        [string]$SiteId,
        [string]$LibraryName,
        [datetime]$ModifiedDate
    )
    
    # Get document library
    $lists = Get-MgSiteList -SiteId $SiteId -Filter "displayName eq '$LibraryName'"
    
    # Get all folders
    $folders = Get-MgSiteListItem -SiteId $SiteId -ListId $list.Id `
        -Filter "folder ne null" -Expand "folder"
    
    # Check each folder for emptiness
    foreach ($folder in $folders) {
        $children = Get-MgSiteListItem -SiteId $SiteId -ListId $list.Id `
            -Filter "parentReference/id eq '$($folder.Id)'"
        
        if ($children.Count -eq 0) {
            # Empty folder found
        }
    }
}
```

#### GUI Design (CleanupGUI.ps1)
- Tab 1: Connection (Site URL, Connect button)
- Tab 2: Scan & Clean (Library dropdown, Date picker, Scan button, Results grid)
- Tab 3: Logs (Activity viewer)
- Status bar with progress
- Preview mode checkbox

### Phase 4: Development Steps

1. **Create new project structure**
   - Clean directory with Graph SDK architecture
   - No dual deployment or backward compatibility

2. **Implement Core Module**
   - GraphAuthentication.ps1 - Handle Connect-MgGraph
   - SharePointOperations.ps1 - Find/delete empty folders
   - Logger.ps1 - Keep existing logging logic

3. **Build GUI**
   - Windows Forms interface
   - Three tabs: Connect, Scan, Logs
   - Progress indicators
   - Export to CSV functionality

4. **Add CLI Support**
   - Command-line parameters
   - Silent mode for automation
   - Export results option

5. **Testing**
   - Authentication test
   - Empty folder detection
   - Deletion operations
   - GUI functionality

### Phase 5: Migration Checklist

- [x] Understand requirements
- [ ] Create CLAUDE.md with project rules
- [ ] Set up Graph SDK project structure
- [ ] Implement authentication module
- [ ] Build folder scanning logic
- [ ] Create deletion functionality
- [ ] Design and build GUI
- [ ] Add CLI interface
- [ ] Implement logging
- [ ] Test all functionality
- [ ] Clean up old PnP code
- [ ] Update documentation
- [ ] Final commit

## Success Criteria
1. Authentication works without app registration
2. GUI launches and is user-friendly
3. Can find empty folders by date
4. Preview mode works (no accidental deletions)
5. Audit logging captures all operations
6. Works across multiple tenants without configuration

## Timeline
- Hour 1: Setup and core authentication
- Hour 2: Folder operations (scan/delete)
- Hour 3: GUI implementation
- Hour 4: Testing and refinement
- Hour 5: Documentation and cleanup

## Notes
- Start completely fresh - no legacy code
- Focus on simplicity and reliability
- Test authentication first before building features
- Keep GUI simple and intuitive
- Ensure comprehensive logging