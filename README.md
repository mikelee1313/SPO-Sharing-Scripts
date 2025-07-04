# SPO-Sharing-Scripts

Used to locate Sharing links and users with access across SharePoint Online / OneDrive Sites.

## Overview

This repository provides PowerShell scripts for Microsoft 365 administrators to audit and manage sharing links, permissions, and user access in SharePoint Online (SPO) and OneDrive for Business.

## Scripts

- [Get-SPOItemPermissions.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Get-SPOItemPermissions.ps1)  
  Retrieves the permissions set on a specific SPO item (file or folder). Useful for detailed item-level access audits.

- [Get-SPOSharingLinks.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Get-SPOSharingLinks.ps1)  
  Finds and lists all sharing links for SPO/OneDrive items, including link types and associated users.

- [Get-SPOSharingLinks-pnp3x.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Get-SPOSharingLinks-pnp3x.ps1)  
  Variant of the above script, adapted for use with PnP PowerShell 3.x.

- [Get-SPSitesAndUsersInfo.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Get-SPSitesAndUsersInfo.ps1)  
  Enumerates all SPO sites and lists users who have access, including sharing details.

- [SPO-RemoveSharedLinks.ps1](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/SPO-RemoveSharedLinks.ps1)  
  Scans for and removes sharing links from SPO items to tighten security.

- [get-allgroups](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/get-allgroups)  
  A script or to enumerate all groups in Entra AD
  
## Documentation

Some scripts have additional documentation:
- [README - Get-SPSitesAndUsersInfo.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/README%20-%20Get-SPSitesAndUsersInfo.md)
- [Readme-Get-SPOItemPermissions.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Readme-Get-SPOItemPermissions.md)
- [Readme-Get-SPOSharingLinks.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Readme-Get-SPOSharingLinks.md)
- [SPO SharingLinks Info.md](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/SPO%20SharingLinks%20Info.md)

## Prerequisites

- PowerShell 5.x or later
- SharePoint Online Management Shell or PnP PowerShell modules
- Appropriate permissions in your Microsoft 365 tenant


## License

MIT License. See [LICENSE](LICENSE) for details.

---

For more scripts and details, view the full directory:  
https://github.com/mikelee1313/SPO-Sharing-Scripts/tree/main/
