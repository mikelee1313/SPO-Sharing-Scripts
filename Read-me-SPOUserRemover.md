# SPOUserRemover.ps1

**Removes specified users from SharePoint Online site collections, including site group memberships, direct file/item permissions, and sharing links.**  
Author: Mike Lee  
Created: 7/29/2025

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [User Configuration](#user-configuration)
- [How It Works](#how-it-works)
- [Usage](#usage)
- [Parameters](#parameters)
- [Logging](#logging)
- [Notes](#notes)
- [Disclaimer](#disclaimer)
- [Support](#support)

---

## Overview

**SPOUserRemover.ps1** is a comprehensive PowerShell script to centrally remove user access from a SharePoint Online site collection, including:

- Site group memberships
- Direct permissions on files and list items
- Access via sharing links and sharing-related groups

The script uses PnP PowerShell modules for robust site interaction and implements intelligent retry/throttling logic to manage SharePoint Online API rate limits. All steps and results are logged for auditing and troubleshooting.

---

## Features

- **Bulk user removal** from groups, items, and sharing links
- **Automated throttling/retry handling** for API calls
- **Detailed logging** of all operations
- **Flexible user input** via text file (one user per line)
- **Multiple layers of SharePoint access removal** (groups, permissions, links)
- **Summary output** at script completion

---

## Prerequisites

- **PnP.PowerShell module** installed ([Docs](https://pnp.github.io/powershell/))
- An **Azure/Entra App Registration** with the following permissions:
    - `Sites.FullControl.All`
- An **X.509 certificate** for app authentication (thumbprint required)
- **User list text file** containing one email/login per line
- PowerShell 7+ recommended

---

## User Configuration

Edit the following variables at the top of the script to match your environment:

```powershell
# Tenant and App Registration Details
$appID = "<Your-Entra-App-Id>"           # Azure App Registration Client ID
$thumbprint = "<Your-Cert-Thumbprint>"   # Certificate Thumbprint
$tenant = "<Your-Tenant-Id>"             # Azure/M365 Tenant ID

# Site and User Configuration
$siteURL = "https://<tenant>.sharepoint.com/sites/<sitename>"     # Site Collection URL
$userListPath = 'C:\temp\UsersList.txt'                          # Path to user list file
```
**User List File Format:**  
Plain text, one email/login per line, e.g.:
```
user1@example.com
user2@example.com
```

---

## How It Works

1. **Connects** to SharePoint Online using certificate-based authentication.
2. **Reads** the user list from the specified file.
3. **Removes users** from all site groups.
4. **Removes direct permissions** on files and list items (including unique permissions).
5. **Revokes sharing links** that grant access to those users.
6. **Cleans up sharing-related groups** and validates removal.
7. **Logs** every action and error to a timestamped log file.
8. **Disconnects** from SharePoint Online and summarizes the results.

---

## Usage

1. **Edit the configuration** variables at the top of `SPOUserRemover.ps1`.
2. **Prepare the user list file** (plain text, one user per line).
3. **Run the script** in PowerShell:
    ```powershell
    .\SPOUserRemover.ps1
    ```
4. **Review the summary** and check the generated log file (path is displayed at completion).

---

## Parameters

All variables are configured at the top of the script.  
The script does not use command-line parameters but can be easily modified to accept them.

| Variable        | Description                                        |
|-----------------|----------------------------------------------------|
| `$appID`        | Azure/Entra Application (Client) ID                |
| `$thumbprint`   | Certificate thumbprint for app authentication      |
| `$tenant`       | Azure/M365 Tenant ID                               |
| `$siteURL`      | Full URL of the target SharePoint site collection  |
| `$userListPath` | Path to the user list file                         |

---

## Logging

- All script actions, successes, warnings, and errors are written to a log file in your temp directory:
    ```
    %TEMP%\SPOUserRemover_<timestamp>.log
    ```
- The log file path is displayed in the console after running the script.

---

## Notes

- The script removes users from **site groups, direct item/file permissions, and sharing links** for maximum coverage.
- Throttling and retry logic helps avoid failures due to SharePoint Online rate limits.
- For large sites, execution may take significant time depending on the number of lists, libraries, and items.

---

## Disclaimer

The sample scripts are provided **AS IS** without warranty of any kind.  
Microsoft further disclaims all implied warranties, including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.  
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you.  
In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

---

**Repository:**  
[https://github.com/mikelee1313/SPO-Sharing-Scripts](https://github.com/mikelee1313/SPO-Sharing-Scripts)
