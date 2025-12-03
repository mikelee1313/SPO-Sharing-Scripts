# SharePoint Document Permission Checker

A PowerShell script that checks if a specific user has access to a SharePoint document and reports how that access is granted (direct permissions, group membership, or "Everyone except external users").

## Features

- ✅ **Accurate Permission Detection**: Handles documents with both inherited and unique (broken inheritance) permissions
- ✅ **Multiple Access Methods**: Detects access via:
  - Direct user permissions
  - SharePoint group membership
  - Microsoft 365/Entra ID group membership
  - "Everyone except external users" permissions
- ✅ **Graph API Validation**: Cross-validates Graph API results with SharePoint REST API roleassignments to filter out stale/cached data
- ✅ **Comprehensive Group Support**: Automatically detects and checks groups that Graph API may miss
- ✅ **Document Owner Information**: Retrieves and displays document owner with email (includes Entra ID lookup fallback)
- ✅ **CSV Export**: Outputs results to CSV for easy analysis

## Prerequisites

### Required Permissions

The Entra ID application must have the following API permissions:

**Microsoft Graph API:**
- `Sites.Read.All` (Application)
- `User.Read.All` (Application)
- `Group.Read.All` (Application)
- `GroupMember.Read.All` (Application)

**SharePoint:**
- `Sites.FullControl.All` (Application)

### Certificate-Based Authentication

This script uses certificate-based authentication with an Entra ID App Registration. You'll need:

1. An Entra ID App Registration
2. A certificate uploaded to the app registration
3. The certificate installed in the Current User certificate store on the machine running the script

## Configuration

Edit the configuration section at the top of the script:

```powershell
# App-Only Authentication Settings
$appID = "your-app-id-here"
$thumbprint = "your-certificate-thumbprint-here"
$tenant = "your-tenant-id-here"

# Tenant Settings
$t = 'YourTenantName'  # Without .onmicrosoft.com

# Single site and user to check
$siteUrl = 'https://yourtenant.sharepoint.com/sites/YourSite'
$userToCheck = 'user@yourtenant.com'

# Document URL
$documentUrl = 'https://yourtenant.sharepoint.com/sites/YourSite/Library/Document.docx'

# Debug mode
$debug = $true  # Set to $false for minimal output
```

## Usage

### Basic Usage

```powershell
.\perms-finder.ps1
```

### Output

The script generates two files in your temp directory:

1. **CSV Report**: `CheckAccess_[timestamp]output.csv`
   - Contains document access details
   - Includes site name, document URL, user, owner, and access type

2. **Log File**: `CheckAccess_[timestamp]logfile.log`
   - Detailed execution log
   - Useful for troubleshooting

### Sample Output

**Console Output:**
```
✓ SUCCESS: Found user@domain.com on document 'Report.docx' with access: Via Entra Group: Marketing (write)
```

**CSV Output:**
| SiteName | URL | DocumentName | DocumentURL | User | Owner | AccessType |
|----------|-----|--------------|-------------|------|-------|------------|
| Marketing | https://tenant.sharepoint.com/sites/Marketing | Report.docx | https://...| user@domain.com | John Doe (john@domain.com) | Via Entra Group: Marketing (write) |

## How It Works

### Permission Detection Flow

1. **Authentication**: Acquires access tokens for both Microsoft Graph API and SharePoint REST API
2. **Site Retrieval**: Locates the SharePoint site
3. **Document Retrieval**: Gets the document using SharePoint REST API (more reliable than Graph API for URL-based retrieval)
4. **Inheritance Check**: Determines if the document has unique permissions or inherits from parent
5. **Permission Analysis**:
   - If **unique permissions**: Queries SharePoint roleassignments for authoritative permission list
   - Validates Graph API results against roleassignments to filter stale data
   - Checks groups that Graph API missed by querying roleassignments
6. **Group Membership**: Checks if user is member of any groups with permissions:
   - SharePoint groups via SharePoint REST API
   - Entra ID groups via Microsoft Graph API
7. **Effective Permissions**: Falls back to SharePoint REST API for effective permission validation
8. **EEEU Validation**: Confirms "Everyone except external users" access and validates user qualifies

### Why SharePoint REST API?

The script uses both Microsoft Graph API and SharePoint REST API because:

- **Graph API** is fast but sometimes returns stale/cached permissions, especially for documents with unique permissions
- **SharePoint REST API** provides authoritative, real-time permission data via roleassignments
- The script cross-validates Graph API results with SharePoint REST API for accuracy

## Access Type Reporting

The script reports access in the following formats:

| Access Type | Description |
|-------------|-------------|
| `Direct Document Access (read/write/owner)` | User has direct permissions on the document |
| `Via SharePoint Group: GroupName (read/write/owner)` | Access through SharePoint group membership |
| `Via Entra Group: GroupName (read/write/owner)` | Access through Microsoft 365/Entra ID group membership |
| `Via M365 Group: GroupName (read/write/owner)` | Access through M365 group membership |
| `Via Everyone Except External Users (read/write)` | Access through "Everyone except external users" permission |
| `No Access Found` | User does not have access to the document |

## Debug Mode

When `$debug = $true`, the script provides detailed output including:

- API calls and responses
- Group membership checks
- Permission validation steps
- Graph API vs SharePoint REST API comparisons
- Inheritance status
- Group detection from roleassignments

Set `$debug = $false` for production use with minimal console output.

## Troubleshooting

### Common Issues

**"Could not find user object"**
- Verify the user email address is correct
- Ensure the app has `User.Read.All` permission

**"Could not retrieve document from URL"**
- Check the document URL is correct and properly encoded
- Verify the document exists and hasn't been moved/deleted
- Ensure URL uses correct library name (e.g., "Shared Documents" not "Documents")

**"Response status code: 404"**
- Document URL may be incorrect
- Library name in URL may need to be "Shared Documents" instead of display name

**"Graph API returned stale data"**
- This is expected behavior for documents with unique permissions
- The script automatically validates against SharePoint REST API
- Access will still be detected correctly via roleassignments

**Permission Denied Errors**
- Verify app registration has all required API permissions
- Ensure permissions are granted admin consent
- Check certificate is valid and installed correctly

## Limitations

- Currently checks one user and one document per execution
- Requires certificate-based authentication (interactive auth not supported)
- Large SharePoint groups (1000+ members) may take time to enumerate

## Technical Details

### API Endpoints Used

**Microsoft Graph API:**
- `/v1.0/sites` - Site retrieval
- `/v1.0/users` - User lookup and group membership
- `/v1.0/groups/{id}/members` - Group member enumeration
- `/v1.0/drives/{id}/items/{id}/permissions` - Document permissions

**SharePoint REST API:**
- `/_api/web/GetFileByServerRelativePath` - Document retrieval by path
- `/_api/web/lists/GetByTitle('{library}')/items({id})/roleassignments` - Authoritative permissions
- `/_api/web/sitegroups/GetById({id})/users` - SharePoint group members
- `/_api/web/GetFileByServerRelativeUrl('{path}')/ListItemAllFields/GetUserEffectivePermissions` - Effective permissions

### Permission Role Mapping

| SharePoint Role | Script Output |
|-----------------|---------------|
| Full Control | owner |
| Edit / Contribute | write |
| Read | read |
| Limited Access | (filtered out - not meaningful) |

## Version History

### Current Version
- Fixed Graph API stale data issues for documents with unique permissions
- Added roleassignments validation for authoritative permission data
- Improved group detection (handles groups Graph API misses)
- Added duplicate EEEU detection prevention
- Enhanced Entra ID group support
- Added document owner retrieval with Entra ID fallback

## Contributing

Issues and pull requests are welcome. Please ensure:
- Code follows existing patterns
- Debug output is helpful and clear
- Changes are tested with both inherited and unique permissions

## License

This script is provided as-is for use in SharePoint permission auditing scenarios.

## Author

Created for SharePoint permission auditing and compliance scenarios.

---

**Note**: This script is designed for SharePoint Online with Entra ID (Azure AD) authentication. It does not support SharePoint On-Premises or SharePoint Server.
