# SPO-Sharing-Scripts

## Overview

**SPO-Sharing-Scripts** is a collection of PowerShell scripts designed for Microsoft 365 administrators to audit, manage, and remediate sharing links, permissions, and user/group access across SharePoint Online and OneDrive sites. These scripts help you inventory sharing links, analyze user permissions, remove users, and export group information, streamlining security and compliance tasks in your tenant.

---

## PowerShell Scripts Included

### [Get-SPOItemPermissions.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Get-SPOItemPermissions.ps1)  
Scans SharePoint Online sites to identify all files and folders and their permissions. Outputs detailed permissions information to Excel, including inheritance and user/group roles.  
- **Documentation:** [Readme-Get-SPOItemPermissions.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Readme-Get-SPOItemPermissions.md)

---

### [Get-SPSitesAndUsersInfo.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Get-SPSitesAndUsersInfo.ps1)  
Collects comprehensive information about SPO sites and users, including site properties, group memberships, direct users, Entra/M365 Group associations, and access details, exporting results to CSV.  
- **Documentation:** [Readme-Get-SPOSitesAndUserInfo.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Readme-Get-SPSitesAndUsersInfo.md)

---

### [Get-and-Remove-SPOSharingLinks.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Get-and-Remove-SPOSharingLinks.ps1)  
Inventories sharing links across SPO sites, identifies Organization and Flexible links, and optionally converts Organization links to direct permissions with cleanup capabilities. Supports both detection (report) and remediation modes.  
- **Documentation:** [Readme-Get-and-Remove-SPOSharingLinks.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Readme-Get-and-Remove-SPOSharingLinks.md)  
- **Additional Info:** [SPO SharingLinks Info.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/SPO%20SharingLinks%20Info.md)

---

### [get-allgroups.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/get-allgroups.ps1)  
Connects to Exchange Online and exports details of unified groups (M365/Entra Groups) filtered by access type (public/private/all) to CSV, including group alias, email, GUID, and associated SharePoint site URLs.

---

### [SPOUserRemover.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/SPOUserRemover.ps1)  
Removes specified users from SPO site collections, targeting group memberships, direct file/item permissions, and sharing links. Includes logging and throttling handling for robust batch operations.

---

## Additional Documentation

- **Get-SPOItemPermissions:** [Readme-Get-SPOItemPermissions.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Readme-Get-SPOItemPermissions.md)
- **Get-SPSitesAndUsersInfo:** [README - Get-SPSitesAndUsersInfo.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/README%20-%20Get-SPSitesAndUsersInfo.md)
- **Get-and-Remove-SPOSharingLinks:** [Readme-Get-and-Remove-SPOSharingLinks.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Readme-Get-and-Remove-SPOSharingLinks.md)
- **General Sharing Links Info:** [SPO SharingLinks Info.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/SPO%20SharingLinks%20Info.md)

---

## Prerequisites

- PowerShell 5.x or later
- SharePoint Online Management Shell or PnP PowerShell modules
- Appropriate permissions in your Microsoft 365 tenant

---

## License

MIT License. See [LICENSE](LICENSE) for details.
