# Get-SPOSharingLinks.ps1

## Overview

The **Get-SPOSharingLinks.ps1** script is a SharePoint Online Sharing Information Collection Tool designed to gather detailed information about SharePoint Online sites and their sharing configurations. It focuses specifically on external sharing links, providing insights into document URLs, users with access, and document ownership. 

This script uses app-only authentication to connect to SharePoint Online and offers the flexibility to process sites from a CSV file or retrieve all sites in the tenant.

---

## Features

- Connects to SharePoint Online using app-only authentication.
- Retrieves site information including Information Barrier settings, sharing configurations, and templates.
- Identifies and extracts sharing links and associated document details.
- Uses Microsoft Graph API to locate documents being shared and their owners.
- Consolidates and exports data into structured CSV files.
- Logs progress and errors in a dedicated log file.

---

## Prerequisites

To run this script, you must have the following:

1. **PnP PowerShell Module**: Install the PnP PowerShell module on your system.
2. **App Registration in Entra ID**:
   - App ID for authentication.
   - Certificate thumbprint for app-only authentication.
   - Tenant ID (GUID) of your Azure Active Directory.
3. **Permissions**:
   - App-only authentication permissions for SharePoint Online and Microsoft Graph API.

---

## Required Configuration

Before running the script, update the following variables in the script:

```powershell
$tenantname = "your-tenant-name"          # Your SharePoint tenant name (without ".sharepoint.com")
$appID = "your-app-id"                    # App ID for Entra ID
$thumbprint = "your-certificate-thumbprint" # Certificate thumbprint
$tenant = "your-tenant-id"                # Tenant ID (GUID)
$searchRegion = "your-region"             # Graph search region (e.g., NAM, EMEA)
```

---

## Usage Instructions

1. **Input File (Optional)**:
   - Provide a CSV file containing site URLs (e.g., `sitelist.csv`) in the format:
     ```
     https://tenant.sharepoint.com/sites/site1
     https://tenant.sharepoint.com/sites/site2
     ```

2. **Run the Script**:
   Execute the script in PowerShell:
   ```powershell
   .\Get-SPOSharingLinks.ps1
   ```

3. **Outputs**:
   - **Main Output CSV**: Contains basic information for all sites.
   - **Sharing Links CSV**: Detailed sharing link data including document URLs, users with access, and document owners.
   - **Log File**: Logs the script's execution progress and errors.

---

## Example Output

### Main Output (Site Information)
| URL                                   | Owner         | Template  | Sharing Capability | Last Content Modified |
|---------------------------------------|---------------|-----------|---------------------|------------------------|
| https://tenant.sharepoint.com/sites/A | user@domain.com | Team Site | Enabled             | 2025-05-01            |

### Sharing Links Output
| Site URL                              | File URL                                      | File Owner         | Sharing Group Name | Sharing Link Members           |
|---------------------------------------|-----------------------------------------------|--------------------|--------------------|---------------------------------|
| https://tenant.sharepoint.com/sites/A | https://tenant-my.sharepoint.com/p/file.docx | John Doe <email>   | SharingLinks12345  | Alice <alice@domain.com>; Bob |

---

## Error Handling

- Errors related to authentication, site retrieval, or data export are logged in the log file.
- The script exits gracefully on critical errors, ensuring partial data is saved whenever possible.

---

## Disclaimer

This script is provided "AS IS" without warranty of any kind. Use it at your own risk. The author and contributors are not liable for any damages arising from its use.

---

## Author

- **Name**: Mike Lee
- **Date Created**: May 5, 2025

For further assistance or inquiries, please contact the repository owner.
