# SharePoint Authentication Strategy for Multi-Tenant Client Tool
*Last Updated: 2025-09-04*

## Current Situation
Our SharePoint Cleanup Tool currently uses PnP.PowerShell v1.12.0 with a workaround using Azure CLI's client ID. This is not ideal for a general-purpose tool that should work across multiple tenants without requiring permanent Azure app registration.

## Authentication Requirements
1. **Must support MFA** (multi-factor authentication)
2. **No permanent Azure app registration** required for end users
3. **Works across multiple tenants** without reconfiguration
4. **Client-side tool** - downloadable and usable by anyone
5. **Simple for non-technical users** (junior admins)

## Available Options Analysis

### Option 1: Microsoft Graph PowerShell SDK (RECOMMENDED)
**Pros:**
- Built-in multi-tenant app registration by Microsoft
- No need for users to register their own app
- Supports MFA and modern authentication
- Works with `Connect-MgGraph` using delegated permissions
- Cross-platform support (Windows, macOS, Linux)
- Actively maintained by Microsoft

**Cons:**
- Requires rewriting from PnP cmdlets to Graph SDK cmdlets
- Different API structure than SharePoint CSOM
- May have limitations for some SharePoint-specific operations

**Implementation:**
```powershell
# Simple connection without app registration
Connect-MgGraph -Scopes "Sites.Read.All", "Sites.ReadWrite.All"

# Access SharePoint sites
$site = Get-MgSite -SiteId "contoso.sharepoint.com:/sites/teamsite"
$drives = Get-MgSiteDrive -SiteId $site.Id
```

### Option 2: SharePoint REST API with MSAL.PS
**Pros:**
- Direct SharePoint API access
- Can use Microsoft's public client applications
- Supports interactive authentication with MFA
- No app registration needed if using public clients

**Cons:**
- Requires handling OAuth tokens manually
- More complex implementation
- Need to manage token refresh

**Implementation:**
```powershell
# Using MSAL.PS module
Install-Module MSAL.PS
$token = Get-MsalToken -ClientId "public-client-id" -TenantId "common" -Interactive
# Use token with REST API calls
```

### Option 3: PnP.PowerShell with Public Client IDs
**Current Workaround:**
- Using Azure CLI client ID: `04b07795-8ddb-461a-bbee-02f9e1bf7b46`
- Works but not officially supported long-term

**Other Public Client IDs to Consider:**
- Microsoft Graph PowerShell: `14d82eec-204b-4c2f-b7e8-296a70dab67e`
- PowerShell Core: `1950a258-227b-4e31-a9cf-717495945fc2`
- Visual Studio: `872cd9fa-d31f-45e0-9eab-6e460a02d1f1`

**Limitations:**
- These client IDs might not have SharePoint permissions by default
- Could be blocked by tenant administrators
- Not guaranteed to work long-term

### Option 4: Hybrid Approach - Multiple Authentication Methods
**Strategy:**
Implement multiple authentication providers and let the tool automatically try them in order:

1. Microsoft Graph SDK (primary)
2. PnP.PowerShell with Azure CLI client ID (fallback)
3. User-provided app registration (advanced users)

## Recommended Migration Path

### Phase 1: Short-term (Current)
- Continue using PnP.PowerShell with Azure CLI client ID workaround
- Document the limitation for users
- Provide option for users to supply their own client ID

### Phase 2: Medium-term (3-6 months)
- Develop parallel implementation using Microsoft Graph PowerShell SDK
- Create abstraction layer to support both PnP and Graph SDK
- Allow users to choose authentication method

### Phase 3: Long-term (6-12 months)
- Fully migrate to Microsoft Graph PowerShell SDK
- Remove dependency on PnP.PowerShell
- Maintain backward compatibility through configuration

## Implementation Considerations

### For Microsoft Graph SDK Migration
1. **Install SDK:**
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```

2. **Key Cmdlet Mappings:**
   - `Get-PnPWeb` → `Get-MgSite`
   - `Get-PnPList` → `Get-MgSiteList`
   - `Get-PnPListItem` → `Get-MgSiteListItem`
   - `Remove-PnPListItem` → `Remove-MgSiteListItem`

3. **Authentication Flow:**
   ```powershell
   # No app registration needed
   Connect-MgGraph -Scopes "Sites.ReadWrite.All" -TenantId "common"
   ```

### For User Experience
1. **First Run Setup:**
   - Auto-detect available authentication methods
   - Guide user through authentication
   - Save preferences for future use

2. **Configuration File:**
   ```json
   {
     "authMethod": "MicrosoftGraph",
     "fallbackMethods": ["PnPAzureCLI", "UserProvided"],
     "customClientId": null,
     "preferredTenantId": "common"
   }
   ```

## Alternative Solutions for Specific Scenarios

### For Enterprise Deployments
- Provide PowerShell script to register app in tenant
- Generate certificate for app-only authentication
- Use managed identity for Azure-hosted scenarios

### For Advanced Users
- Allow custom authentication providers
- Support certificate-based authentication
- Enable app-only access for automation

## Security Considerations
1. **Never store credentials** in the tool
2. **Use system credential manager** for token caching
3. **Implement proper token refresh** logic
4. **Support conditional access** policies
5. **Log authentication attempts** for audit

## Decision Matrix

| Criteria | PnP.PowerShell | MS Graph SDK | REST API | Hybrid |
|----------|---------------|--------------|----------|--------|
| No App Registration | ❌* | ✅ | ❌ | ✅ |
| MFA Support | ✅ | ✅ | ✅ | ✅ |
| Multi-tenant | ⚠️ | ✅ | ✅ | ✅ |
| Ease of Implementation | ✅ | ⚠️ | ❌ | ❌ |
| Future Proof | ❌ | ✅ | ⚠️ | ✅ |
| SharePoint Features | ✅ | ⚠️ | ✅ | ✅ |

*Currently using workaround with Azure CLI client ID

## Final Recommendation
**Migrate to Microsoft Graph PowerShell SDK** as the primary authentication and API method because:
1. It has built-in multi-tenant support without app registration
2. It's actively maintained by Microsoft
3. It supports all required authentication scenarios
4. It's the strategic direction for Microsoft 365 development

Keep PnP.PowerShell as a fallback option during transition period to ensure compatibility for users who cannot immediately migrate.

## Action Items
- [ ] Create proof-of-concept using Microsoft Graph SDK
- [ ] Test folder deletion operations via Graph API
- [ ] Benchmark performance vs PnP.PowerShell
- [ ] Design abstraction layer for authentication
- [ ] Create migration guide for users
- [ ] Update documentation with new authentication flow

## Resources
- [Microsoft Graph PowerShell SDK Documentation](https://learn.microsoft.com/en-us/powershell/microsoftgraph/)
- [PnP.PowerShell Authentication Guide](https://pnp.github.io/powershell/articles/authentication.html)
- [SharePoint REST API Reference](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [MSAL.PS Module](https://github.com/AzureAD/MSAL.PS)