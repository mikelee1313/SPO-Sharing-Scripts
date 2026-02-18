# SPOUserRemover.ps1

A PowerShell script for bulk removal of user access from SharePoint Online site collections. It systematically removes users from site groups, direct file/item permissions, sharing links, and optionally the User Information List (UIL).

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Setup](#setup)
  - [1. Azure/Entra App Registration](#1-azureentra-app-registration)
  - [2. Certificate Authentication](#2-certificate-authentication)
  - [3. Install PnP PowerShell](#3-install-pnp-powershell)
  - [4. Prepare Input Files](#4-prepare-input-files)
- [Configuration](#configuration)
- [Usage](#usage)
  - [Single Site Mode](#single-site-mode)
  - [Multi-Site Mode](#multi-site-mode)
- [How It Works](#how-it-works)
  - [Phase 1 — Site Group Removal](#phase-1--site-group-removal)
  - [Phase 2 — File & Item Permission Removal](#phase-2--file--item-permission-removal)
  - [Phase 3 — Sharing Link Access Removal](#phase-3--sharing-link-access-removal)
  - [Phase 4 — User Information List Removal (Optional)](#phase-4--user-information-list-removal-optional)
- [Throttle Handling](#throttle-handling)
- [Logging](#logging)
- [Important Notes](#important-notes)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)
- [Disclaimer](#disclaimer)

## Overview

When offboarding users or revoking access across SharePoint Online site collections, manually removing permissions is time-consuming and error-prone. **SPOUserRemover** automates this process by:

1. Removing users from all **site groups** (Members, Owners, Visitors, custom groups)
2. Removing **direct permissions** on individual files and list items with unique role assignments
3. Removing users from **sharing-related groups** without revoking the sharing links themselves (preserving access for other legitimate users)
4. Optionally removing users from the **User Information List** to fully clean up their presence in the site

The script supports processing a single site or iterating over a list of multiple site collections.

## Features

- **Multi-site support** — Process one site or hundreds from a text file
- **Comprehensive permission removal** — Covers groups, direct file/item permissions, and sharing links
- **Non-destructive sharing link handling** — Removes target users from sharing groups without revoking links for other users
- **Intelligent throttle handling** — Automatic retry with exponential backoff on HTTP 429 responses
- **Detailed logging** — Timestamped log file for every operation, written to `%TEMP%`
- **User Information List cleanup** — Optionally removes users from the hidden UIL
- **Certificate-based authentication** — Secure, non-interactive authentication via Azure/Entra app registration

## Prerequisites

| Requirement | Details |
|---|---|
| **PowerShell** | 5.1+ (Windows PowerShell) or 7.x (PowerShell Core) |
| **PnP.PowerShell** | `PnP.PowerShell` module ([GitHub](https://github.com/pnp/powershell)) |
| **Azure/Entra App Registration** | With **Sites.FullControl.All** application permission (admin-consented) |
| **Certificate** | Uploaded to the app registration and installed in the local certificate store |
| **Permissions** | The app must have admin consent for the `Sites.FullControl.All` Graph/SharePoint permission |

## Setup

### 1. Azure/Entra App Registration

1. Navigate to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) and create a new registration (or use an existing one).
2. Under **API permissions**, add:
   - **SharePoint → Application permissions → Sites.FullControl.All**
3. Grant **admin consent** for the permission.
4. Note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

### 2. Certificate Authentication

1. Create or obtain a certificate (self-signed is fine for internal use):
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=SPOUserRemover" `
       -CertStoreLocation "Cert:\CurrentUser\My" `
       -KeyExportPolicy Exportable -KeySpec Signature `
       -KeyLength 2048 -NotAfter (Get-Date).AddYears(2)
   ```
2. Export the public key (`.cer`) and upload it to the app registration under **Certificates & secrets → Certificates**.
3. Note the **certificate thumbprint**.

### 3. Install PnP PowerShell

```powershell
Install-Module -Name PnP.PowerShell -Scope CurrentUser
```

### 4. Prepare Input Files

#### User List File (`UsersList.txt`)

A plain text file with one user email or login name per line:

```
jdoe@contoso.com
asmith@contoso.com
bwilson@contoso.com
```

#### Site List File (`SiteList.txt`) — *only needed for multi-site mode*

A plain text file with one SharePoint site URL per line. Lines starting with `#` are treated as comments and ignored:

```
https://contoso.sharepoint.com/sites/ProjectAlpha
https://contoso.sharepoint.com/sites/ProjectBeta
# https://contoso.sharepoint.com/sites/Archived  ← skipped
https://contoso.sharepoint.com/sites/TeamSite01
```

## Configuration

Open the script and update the **USER CONFIGURATION** section at the top:

```powershell
# --- Tenant and App Registration Details ---
$appID      = "<your-app-client-id>"          # Entra App (Client) ID
$thumbprint = "<your-certificate-thumbprint>" # Certificate thumbprint
$tenant     = "<your-tenant-id>"              # Azure/Microsoft 365 Tenant ID

# --- Site and User Configuration ---
$useSiteList   = $true                              # $true = multi-site mode, $false = single site
$siteURL       = "https://contoso.sharepoint.com/sites/MySite"  # Used when $useSiteList = $false
$siteListPath  = "C:\temp\SiteList.txt"             # Path to site list file (used when $useSiteList = $true)
$userListPath  = "C:\temp\UsersList.txt"            # Path to user list file
$RemoveFromUIL = $true                              # $true = also remove from User Information List
```

| Parameter | Type | Description |
|---|---|---|
| `$appID` | String | Azure/Entra application (client) ID |
| `$thumbprint` | String | Certificate thumbprint for authentication |
| `$tenant` | String | Microsoft 365 tenant ID |
| `$useSiteList` | Boolean | `$true` to process multiple sites from a file; `$false` for a single site |
| `$siteURL` | String | SharePoint site URL (used only when `$useSiteList` is `$false`) |
| `$siteListPath` | String | Path to a text file containing site URLs (used only when `$useSiteList` is `$true`) |
| `$userListPath` | String | Path to a text file containing user emails/login names |
| `$RemoveFromUIL` | Boolean | `$true` to remove users from the User Information List after permission removal |

## Usage

### Single Site Mode

1. Set `$useSiteList = $false`
2. Set `$siteURL` to the target site collection URL
3. Run the script:

```powershell
.\SPOUserRemover.ps1
```

### Multi-Site Mode

1. Set `$useSiteList = $true`
2. Set `$siteListPath` to the path of your site list file
3. Run the script:

```powershell
.\SPOUserRemover.ps1
```

The script will iterate through each site, connect, process all removals, disconnect, and move to the next site. A summary is displayed at the end.

## How It Works

### Phase 1 — Site Group Removal

- Enumerates all site groups (Members, Owners, Visitors, custom groups)
- For each group, retrieves the member list
- Matches target users by **email**, **login name**, or **partial login name**
- Removes matched users from the group

### Phase 2 — File & Item Permission Removal

- Enumerates all non-hidden document libraries and lists
- For each item, checks if it has **unique role assignments** (broken inheritance)
- If unique permissions exist, loads the role assignments and matches target users
- Removes all non-"Limited Access" roles for matched users on that item
- Handles both document libraries (files) and generic lists (items)

### Phase 3 — Sharing Link Access Removal

This phase uses a **non-destructive approach** — sharing links are **preserved** for legitimate users. Target user access is revoked by removing them from the underlying sharing groups.

**Step 1 — Identification:** Scans document libraries for sharing links and identifies flexible links containing target users.

**Step 2 — Group Cleanup:** Removes target users from sharing-related groups including:
- `SharingLinks.*` groups
- `Everyone except external users`
- Anonymous sharing groups
- SPO grid all-users groups
- Federated directory claim provider groups

**Step 3 — Verification:** Re-scans sharing links to verify that access control has been applied through group membership removal. Also performs a secondary pass on individual file sharing permissions (View Only, Edit, Read roles).

### Phase 4 — User Information List Removal (Optional)

When `$RemoveFromUIL = $true`:
- Looks up each user in the site's User Information List via `Get-PnPUser`
- Removes the user entry with `Remove-PnPUser`, fully cleaning up their presence

## Throttle Handling

SharePoint Online enforces rate limits (HTTP 429). The script handles throttling automatically:

- **Max retries:** 3 per operation
- **Backoff strategy:** Exponential — 5s, 10s, 20s
- All throttle events are logged with timestamps
- Operations resume automatically after the wait period

## Logging

Every run generates a timestamped log file at:

```
%TEMP%\SPOUserRemover_<yyyyMMdd_HHmmss>.log
```

Log entries include:
- Timestamp, severity level (`INFO`, `WARNING`, `ERROR`, `SUCCESS`), and message
- Connection attempts and validation
- Every group membership change
- Every file/item permission removal
- Throttle events and retries
- Summary statistics

The log file path is displayed at the start and end of each run.

## Important Notes

> **Sharing links are NOT revoked.** The script removes target users from sharing groups to revoke their access while preserving links for other authorized users.

- **"Limited Access" roles are skipped** during file/item permission removal, as these are typically system-managed and cannot be removed directly.
- **The script requires `Sites.FullControl.All`** — this is a high-privilege permission. Restrict the app registration to only what is needed and audit usage.
- **Test in a non-production environment first.** Run against a test site collection before processing production sites.
- **Large sites may take significant time** due to the need to enumerate all items and check unique permissions.
- **Users already removed** are handled gracefully — "principal not found" errors are caught and logged as informational messages.

## Examples

### Remove 3 users from a single team site

**UsersList.txt:**
```
john.doe@contoso.com
jane.smith@contoso.com
bob.wilson@contoso.com
```

**Script configuration:**
```powershell
$useSiteList   = $false
$siteURL       = "https://contoso.sharepoint.com/sites/ProjectTeam"
$userListPath  = "C:\temp\UsersList.txt"
$RemoveFromUIL = $true
```

**Run:**
```powershell
.\SPOUserRemover.ps1
```

### Remove users from 50 site collections

**SiteList.txt:**
```
https://contoso.sharepoint.com/sites/Site01
https://contoso.sharepoint.com/sites/Site02
...
https://contoso.sharepoint.com/sites/Site50
```

**Script configuration:**
```powershell
$useSiteList   = $true
$siteListPath  = "C:\temp\SiteList.txt"
$userListPath  = "C:\temp\UsersList.txt"
$RemoveFromUIL = $true
```

**Sample console output:**
```
SharePoint Online User Remover Script
=====================================
Multi-Site Mode: Enabled
Site List: C:\temp\SiteList.txt
User List: C:\temp\UsersList.txt

========================================
Site 1 of 50
========================================
Processing Site: https://contoso.sharepoint.com/sites/Site01
Connecting to SharePoint Online...
Successfully connected to SharePoint Online
Removing users from site groups...
  ...
Completed processing for site: https://contoso.sharepoint.com/sites/Site01

========================================
All Sites Processing Complete
========================================
Total sites processed: 50
Successful: 48
Failed: 2
```

## Troubleshooting

| Issue | Cause | Solution |
|---|---|---|
| `Failed to connect to SharePoint Online` | Incorrect app ID, thumbprint, tenant ID, or site URL | Double-check all four values; ensure the certificate is in the local store |
| `User list file not found` | Incorrect path in `$userListPath` | Verify the file exists at the specified path |
| `Site list file not found` | Incorrect path in `$siteListPath` | Verify the file exists at the specified path |
| `Max retries exceeded` | Excessive throttling from SharePoint Online | Wait and retry later; reduce the number of sites/users per run |
| `Access denied` / `403 Forbidden` | App lacks `Sites.FullControl.All` or admin consent | Re-check API permissions and grant admin consent |
| `Can not find the principal` | User has already been removed | Handled gracefully — informational log, no action needed |
| Script runs slowly on large sites | Every item is checked for unique permissions | Expected behavior; consider running during off-peak hours |

## Disclaimer

The sample scripts are provided **AS IS** without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

## Author

**Mike Lee**

---

*Last updated: July 2025*
