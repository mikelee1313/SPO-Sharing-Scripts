# SPO-OTPHelper.ps1

A PowerShell script to assess the impact of **MC1243549 – Retirement of SharePoint One-Time Passcode (OTP)** and transition to Microsoft Entra B2B guest accounts.

The script scans SharePoint Online sites to identify **Flexible sharing links** that contain external OTP users. Only links that are within scope of the OTP retirement (i.e. links with confirmed external OTP users) are written to the output CSV — giving admins a targeted action list rather than a full inventory dump.

> **Detection-only mode.** This script makes **no changes** to sharing links, permissions, or user accounts.

---

## Table of Contents

- [Background](#background)
- [Prerequisites](#prerequisites)
- [Configuration](#configuration)
- [Usage](#usage)
- [Output](#output)
- [How It Works](#how-it-works)
- [Disclaimer](#disclaimer)

---

## Background

Microsoft Message Center notification **MC1243549** announces the retirement of SharePoint OTP (One-Time Passcode) external sharing and the migration of those users to Microsoft Entra B2B guest accounts.

OTP users are identified by the `urn:spo:guest#` pattern in their SharePoint login name. This script helps admins:

1. Identify which Flexible sharing links expose OTP users.
2. Confirm OTP status by cross-referencing each site's User Information List.
3. Quantify the retirement impact across the entire tenant (or a targeted site list).

---

## Prerequisites

### PowerShell Module

- **PnP.PowerShell 2.x or later**

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### Entra App Registration

Create an app registration in Entra ID (Azure AD) with **certificate-based authentication** and grant the following API permissions:

| API | Permission | Type |
|-----|-----------|------|
| SharePoint | `Sites.FullControl.All` | Application |
| SharePoint | `User.Read.All` | Application |
| Microsoft Graph | `Sites.FullControl.All` | Application |
| Microsoft Graph | `Sites.Read.All` | Application |
| Microsoft Graph | `Files.Read.All` | Application |

Upload or generate a certificate for the app registration and note the **thumbprint**.

---

## Configuration

Open the script and set the variables in the **Set Variables** section at the top:

```powershell
$tenantname      = "contoso"                                    # Tenant name (without .onmicrosoft.com)
$appID           = "00000000-0000-0000-0000-000000000000"       # Entra App (client) ID
$thumbprint      = "AABBCCDDEEFF..."                            # Certificate thumbprint
$tenant          = "00000000-0000-0000-0000-000000000000"       # Tenant ID (GUID)
$searchRegion    = ""                                           # Leave empty to auto-detect, or set explicitly: NAM, EUR, APC, GBR, CAN, IND, AUS, etc.
$debugLogging    = $false                                       # Set $true for verbose debug logging
$GetOneDriveInfo = $false                                       # Set $true to scan OneDrive sites ONLY; $false scans SharePoint and skips OneDrive
```

### Optional: Target specific sites

To scan a specific list of sites instead of the entire tenant, set `$inputfile` to the path of a CSV file:

```powershell
$inputfile = "C:\temp\sites.csv"
```

The CSV can be a plain list of URLs (no header) or include a `URL` header column. If `$inputfile` is empty, all active non-archived sites in the tenant are scanned.

### OneDrive for Business sites

By default (`$GetOneDriveInfo = $false`) the script scans **SharePoint sites only** and skips all OneDrive for Business personal sites. Since a significant volume of external OTP sharing also occurs on OneDrive, you can flip this flag to target OneDrive instead:

```powershell
$GetOneDriveInfo = $true   # Scan OneDrive personal sites ONLY (skips all SharePoint sites)
```

Run the script twice — once with `$false` for SharePoint sites and once with `$true` for OneDrive sites — to get a complete tenant-wide picture.

---

## Usage

```powershell
# Scan all sites in the tenant
.\SPO-OTPHelper.ps1

# Scan a specific list of sites (set $inputfile inside the script first)
.\SPO-OTPHelper.ps1
```

Output files are written to `$env:TEMP` and the paths are printed to the console at the end of the run.

---

## Output

### CSV — `SPO_OTP_Impact_<timestamp>.csv`

Contains one row per Flexible sharing link that has at least one confirmed external OTP user. Links with no OTP users are excluded — they are not impacted by the OTP retirement.

| Column | Description |
|--------|-------------|
| Site URL | SharePoint site containing the sharing link |
| Site Owner | Owner of the SharePoint site |
| IB Mode | Information Barrier mode setting |
| IB Segment | Information Barrier segments |
| Site Template | SharePoint site template |
| Sharing Group Name | Name of the internal SharePoint sharing group |
| Sharing Link Members | Users with access via the sharing link |
| File URL | Direct URL to the shared file or list item |
| File Owner | Owner/creator of the shared file |
| Filename | Name of the shared file or list item |
| SharingType | `Flexible` (Organization links are excluded from this report) |
| Sharing Link URL | Clickable sharing link URL |
| Link Expiration Date | When the sharing link expires (if set) |
| IsTeamsConnected | Whether the site is connected to Microsoft Teams |
| SharingCapability | Site-level sharing capability setting |
| Last Content Modified | Last modification date of the site content |
| Search Status | How the file was located (see values below) |
| Has External OTP Users | `True` for all rows in this report |
| External OTP Users | Semicolon-separated list of external user emails confirmed as OTP users via the site's User Information List |
| OTP Confirmed | `True` if at least one user was verified via the User Information List |

#### Search Status values

| Value | Meaning |
|-------|---------|
| `Found` | Document located via Microsoft Graph search |
| `Found (REST Fallback)` | Document located via SharePoint REST API (not yet indexed in Graph search) |
| `File Not Found` | Document not found via any method — the sharing link likely points to a deleted or moved file (orphaned sharing group) |
| `Search Error` | An unexpected error occurred during all lookup attempts |
| `Not Searched` | Search was not attempted |

### Log — `SPOSharingLinks<timestamp>_logfile.log`

Contains INFO and ERROR entries for every site processed. Set `$debugLogging = $true` for additional verbose DEBUG entries useful for troubleshooting.

---

## How It Works

```
For each site in the tenant (or input CSV):
│
├─ Get all SharePoint groups
│   └─ For each SharingLinks.* group:
│       ├─ Extract the document UniqueId (GUID) from the group name
│       ├─ Search for the document via Microsoft Graph (driveItem → listItem)
│       └─ If not found in Graph → SharePoint REST fallback (4 attempts):
│           1. GetFileById            (document library files)
│           2. SP REST search API     (site-scoped, bypasses Graph restrictions)
│           3. GetListItemByUniqueId  (list items / wiki pages)
│           4. Group Description path (server-relative path set by SharePoint)
│
├─ For each Flexible sharing link:
│   ├─ Retrieve sharing link URL and expiration from Get-PnPFileSharingLink
│   └─ Detect external OTP users:
│       ├─ Identify users matching urn:spo:guest# login pattern
│       └─ Confirm via site's User Information List (Get-PnPUser)
│
└─ Write impacted links to CSV
    └─ Only rows where Has External OTP Users = True are written
```

### Why a REST fallback?

Microsoft Graph search can be restricted by tenant-level search settings. The SharePoint REST search endpoint runs under the PnP certificate connection and is not subject to the same Graph restrictions, allowing the script to locate files that Graph search cannot reach.

---

## Disclaimer

The sample scripts are provided **AS IS** without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

---

*Author: Mike Lee | Updated: March 26, 2026 | MC1243549 – SharePoint OTP Retirement Impact Assessment*
