Certainly! To create a README for the script Readme-Get-SPOItemPermissions.md in your repo mikelee1313/SPO-Sharing-Scripts, I need to know what the script does or what its main features/usage are. Since the script is not shown here, I’ll assume it is a PowerShell script for retrieving item permissions in SharePoint Online (SPO) and OneDrive, consistent with your repo description. Here’s a template README you can further customize:

---

# Get-SPOItemPermissions

**Get-SPOItemPermissions** is a PowerShell script designed to help administrators and site owners retrieve and analyze item-level permissions across SharePoint Online and OneDrive for Business. This utility is especially useful for auditing, compliance, and understanding who has access to specific files, folders, or items within your tenant.

## Features

- List user and group permissions for specific items, folders, or document libraries
- Output results to console or export to CSV for reporting
- Supports enumeration of sharing links and unique permissions
- Can be used for compliance checks and permission reviews

## Prerequisites

- PowerShell 5.1 or later (Windows) or PowerShell Core (cross-platform)
- [SharePoint Online Management Shell](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online) or [PnP PowerShell](https://pnp.github.io/powershell/)
- Appropriate permissions to read site and item permissions on the target SharePoint Online or OneDrive sites

## Usage

### Connect to SharePoint Online

```powershell
Connect-SPOService -Url https://<your-tenant>-admin.sharepoint.com
```

### Run the Script

```powershell
.\Get-SPOItemPermissions.ps1 -SiteUrl "https://<your-tenant>.sharepoint.com/sites/yoursite" -ItemUrl "/sites/yoursite/Shared Documents/YourFolder/YourFile.docx"
```

#### Parameters

- `-SiteUrl` (Required): The URL of the SharePoint Online site.
- `-ItemUrl` (Required): The server-relative URL of the item (file/folder) to audit.
- `-ExportCsv` (Optional): Path to export the results as a CSV file.

### Example

```powershell
.\Get-SPOItemPermissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/HR" -ItemUrl "/sites/HR/Shared Documents/Policies/LeavePolicy.docx" -ExportCsv "C:\Reports\LeavePolicyPermissions.csv"
```

## Output

- Console display of users, groups, sharing links, and their permissions
- Optional CSV export for further analysis

## Limitations

- Requires appropriate admin or site owner permissions to query permissions
- May not list permissions inherited from parent objects unless specified

## License

MIT License

---

**Feel free to further customize this README with more specifics about what your script does, its parameters, and any other usage notes! If you paste the script or share its key features, I can tailor this even more closely to your implementation.**
