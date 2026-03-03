# SPO-UserCleanup.ps1

A PowerShell script for auditing and removing SharePoint Online user permissions at scale. It identifies every access path a user has across your tenant — direct permissions, group memberships, sharing links, and more — and can remove them all in a single run.

Originally developed to solve **PUID mismatch issues** caused by former employees whose SharePoint access was never fully cleaned up before rehire.

---

## Table of Contents

- [Overview](#overview)
- [Use Cases](#use-cases)
- [How It Works](#how-it-works)
  - [Report Mode](#report-mode)
  - [Remove Mode](#remove-mode)
  - [Both Mode](#both-mode)
- [Prerequisites](#prerequisites)
  - [PowerShell Module](#powershell-module)
  - [Entra ID App Registration](#entra-id-app-registration)
  - [Certificate Setup](#certificate-setup)
- [Configuration](#configuration)
- [Usage](#usage)
- [Output Files](#output-files)
- [Access Vectors Checked](#access-vectors-checked)
- [What Gets Removed](#what-gets-removed)
- [Throttling Protection](#throttling-protection)
- [UIL Entries Explained](#uil-entries-explained)
- [Flexible Sharing Links](#flexible-sharing-links)
- [Debug Mode](#debug-mode)
- [Notes and Caveats](#notes-and-caveats)

---

## Overview

`SPO-UserCleanup.ps1` is a three-mode tool:

| Mode | Description |
|------|-------------|
| `Report` | Scans all (or a single targeted) SPO site(s), produces a CSV audit report of every site where each user was found and how they have access. |
| `Remove` | Reads a previously generated CSV and removes each user from only the sites where they were found. |
| `Both` | Performs the Report scan and removes users **inline per site** as each site is processed — no separate pass needed. |

---

## Use Cases

- **Pre-rehire cleanup** — Remove all SPO access for a departed employee before they are rehired to prevent PUID mismatch issues.
- **Offboarding audit** — Generate a full access report before or after disabling an account.
- **Compliance review** — Produce a timestamped CSV showing exactly where a user had access and via what mechanism.
- **Bulk user cleanup** — Process a list of multiple users across all sites in a single execution.

---

## How It Works

### Report Mode

1. Connects to the SharePoint admin center using app-only authentication.
2. Enumerates all sites in the tenant (or a single target site).
3. For each site and each user, checks every access vector (see [Access Vectors Checked](#access-vectors-checked)).
4. Writes results incrementally to a timestamped CSV as each site completes.
5. Also writes a timestamped log file.

### Remove Mode

1. Reads the CSV produced by a prior Report run.
2. Groups rows by site URL so each site is processed once with all affected users batched together.
3. For each site: removes from SharePoint groups, direct file/item permissions, sharing link groups, and optionally the User Information List (UIL).

### Both Mode

Combines Report and Remove in a single execution. After completing the scan for each site, users found on that site are **immediately removed inline** — no intermediate CSV round-trip is required. A CSV is still written as an audit record of what was found and removed.

---

## Prerequisites

### PowerShell Module

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

Requires PnP.PowerShell. Tested with PowerShell 7+. Also works with Windows PowerShell 5.1.

### Entra ID App Registration

Create an App Registration in Entra ID (Azure AD) with the following **Application** (not Delegated) API permissions, then grant admin consent:

| API | Permission | Type |
|-----|-----------|------|
| SharePoint | `Sites.FullControl.All` | Application |
| Microsoft Graph | `Sites.Read.All` | Application |

> **Note:** `Sites.FullControl.All` is required for the Remove phase. `Sites.Read.All` (Graph) is used for Entra ID group membership lookups.

### Certificate Setup

The script uses **certificate-based app-only authentication**. You will need:

1. A self-signed or CA-issued certificate.
2. The certificate uploaded to your Entra ID App Registration (under **Certificates & secrets**).
3. The certificate installed in the **local machine or current user certificate store** on the machine running the script.
4. The certificate **thumbprint**.

To create a self-signed certificate with PnP PowerShell:

```powershell
New-PnPAzureCertificate -OutPfx "SPOCleanup.pfx" -OutCert "SPOCleanup.cer" -CertificatePassword (ConvertTo-SecureString -String "YourPassword" -AsPlainText -Force)
```

Upload `SPOCleanup.cer` to the App Registration. Import `SPOCleanup.pfx` to your certificate store.

---

## Configuration

All configuration is in the **USER CONFIGURATION** section at the top of the script. Update these values before running.

```powershell
# Entra ID App Registration
$appID      = "<Your App/Client ID>"
$thumbprint = "<Certificate Thumbprint>"
$tenant     = "<Tenant ID GUID>"
$t          = "<Tenant name>"        # e.g. 'contoso' (no .onmicrosoft.com)
```

```powershell
# Mode: "Report", "Remove", or "Both"
$Mode = "Report"
```

### Report Mode Settings

| Variable | Description |
|----------|-------------|
| `$UsersFilePath` | Path to a plain-text file with one UPN per line (e.g., `C:\temp\users.txt`). |
| `$TargetSiteUrl` | Leave empty to scan **all** sites. Set to a full site URL to target a single site. |
| `$IncludeOneDrive` | `$true` to include OneDrive for Business sites in the scan. Default: `$false`. |
| `$checkEEEU` | `$true` to check "Everyone except external users" permissions. Adds processing time. Default: `$false`. |

### Remove Mode Settings

| Variable | Description |
|----------|-------------|
| `$InputCsvPath` | Full path to the CSV produced by a prior Report run. Auto-set when using `Both` mode. |
| `$RemoveFromUIL` | `$true` to also remove users from the Site User Information List. Default: `$true`. |

### Shared / Output Settings

| Variable | Default | Description |
|----------|---------|-------------|
| `$debug` | `$false` | Set to `$true` for verbose console and log output. |
| `$enableThrottlingProtection` | `$true` | Enables automatic retry with exponential backoff on HTTP 429/503. Recommended for large tenants. |
| `$baseDelayBetweenSites` | `2` | Seconds to pause between sites. |
| `$baseDelayBetweenUsers` | `1` | Seconds to pause between users within a site. |
| `$maxRetryAttempts` | `5` | Maximum retry attempts on throttling. |
| `$baseRetryDelay` | `30` | Base delay in seconds for retry backoff. |

---

## Usage

### Prepare users.txt

Create a plain text file with one UPN per line:

```
john.doe@contoso.com
jane.smith@contoso.com
```

### Run Report Only

```powershell
$Mode = "Report"
$UsersFilePath = "C:\temp\users.txt"
$TargetSiteUrl = ""   # leave empty for all sites
.\SPO-UserCleanup.ps1
```

### Run Remove from Existing CSV

```powershell
$Mode = "Remove"
$InputCsvPath = "C:\Users\you\AppData\Local\Temp\SiteUsers_2026-03-03_10-00-00_output.csv"
.\SPO-UserCleanup.ps1
```

### Run Both (Scan and Remove in One Pass)

```powershell
$Mode = "Both"
$UsersFilePath = "C:\temp\users.txt"
.\SPO-UserCleanup.ps1
```

### Target a Single Site

```powershell
$Mode = "Report"
$UsersFilePath = "C:\temp\users.txt"
$TargetSiteUrl = "https://contoso.sharepoint.com/sites/HR"
.\SPO-UserCleanup.ps1
```

---

## Output Files

Both files are written to `$env:TEMP` with a timestamp in the filename.

| File | Example Name | Description |
|------|-------------|-------------|
| CSV Report | `SiteUsers_2026-03-03_10-00-00_output.csv` | Audit report of all found user access. |
| Log File | `SiteUsers_2026-03-03_10-00-00_logfile.log` | Timestamped log of all operations, errors, and removal actions. |

### CSV Columns

| Column | Description |
|--------|-------------|
| `SiteName` | Display name of the SharePoint site. |
| `URL` | Full URL of the site. |
| `User` | UPN of the user found. |
| `Owner` | Site primary owner. |
| `AccessType` | Semicolon-delimited list of how the user has access (see below). |

### AccessType Examples

```
Direct Access: Edit
SharePoint Group: Contoso Members; M365 Group Member: Project X
Entra ID Group: Security-SPO-ReadAll (Read)
UIL Entry (residual)
UIL Entry (residual – EEEU also active, verify not phantom)
```

---

## Access Vectors Checked

The Report phase checks the following in order for every user/site combination:

| Check | Description |
|-------|-------------|
| **1. Direct permissions** | User has a direct role assignment at the site (web) level. Reports the specific permission level(s). |
| **1.5. M365 Group membership** | Site is connected to a Microsoft 365 Group and user is a member or owner of that group. |
| **2. Site Owner** | User is the primary site owner. |
| **3. Site Collection Administrator** | User is listed as a Site Collection Admin, including via group claims. |
| **4. SharePoint Group membership** | User is a member of any SharePoint group on the site. |
| **5. Entra ID / Azure AD security group** | User is a member of a security group that has been granted permissions on the site. Resolves both tenant claim (`c:0t.c|tenant|`) and federated claim (`c:0o.c|federateddirectoryclaimprovider|`) formats. |
| **6. Everyone except external users** (optional) | Checks if EEEU has active non-limited permissions and the user is an internal tenant member. Controlled by `$checkEEEU`. |
| **7. User Information List fallback** | If no explicit permissions are found, checks the UIL for a residual or phantom entry and qualifies the finding (see [UIL Entries Explained](#uil-entries-explained)). |

---

## What Gets Removed

The Remove phase targets all of the following per site:

| Target | Details |
|--------|---------|
| **SharePoint group memberships** | Removes the user from all SP groups on the site (excluding SharingLinks groups, which are handled separately). |
| **Direct file/item permissions** | Scans all document libraries and lists for items with unique permissions and removes the user's role assignments. Skips items that inherit permissions. |
| **Sharing link groups** | Removes the user from all `SharingLinks.*` SP groups. For **Flexible (specific-people) links**, uses the SharePoint `ShareLink` REST API with `inviteesToRemove` to remove both the SP group membership and the invitation metadata shown in the UI. |
| **User Information List** (optional) | Removes the user's entry from the site's UIL via `Remove-PnPUser`. Controlled by `$RemoveFromUIL`. |

---

## Throttling Protection

When `$enableThrottlingProtection = $true` (default), the script:

- Adds configurable delays between sites and between users.
- Automatically detects HTTP 429 (Too Many Requests) and 503 (Server Too Busy) responses.
- Honors the `Retry-After` header when present.
- Falls back to exponential backoff with jitter when no `Retry-After` is provided.
- Retries up to `$maxRetryAttempts` times before giving up and logging the failure.

This is strongly recommended for tenants with hundreds or thousands of sites.

---

## UIL Entries Explained

The SharePoint **User Information List (UIL)** is a hidden list that records every user who has ever interacted with a site. Entries can persist even after all explicit permissions are removed. The script distinguishes two cases:

| AccessType Value | Meaning |
|-----------------|---------|
| `UIL Entry (residual)` | User is in the UIL, has no current rights, and EEEU is **not** active. Almost certainly a genuine leftover entry safe to clean up. |
| `UIL Entry (residual – EEEU also active, verify not phantom)` | User is in the UIL, has no current rights, but EEEU **is** active on the site. Could be a residual entry **or** a phantom entry auto-created when the user browsed the site via EEEU. Verify intent before removing. |

Setting `$RemoveFromUIL = $true` will remove UIL entries for found users. Set it to `$false` if you want to skip UIL cleanup (e.g., for audit-only runs or when UIL entries are expected).

---

## Flexible Sharing Links

Standard `Remove-PnPGroupMember` only removes the user from the SharePoint group backing the link — it does **not** remove the invitation metadata that the SharePoint UI reads. This causes removed users to still appear in the "Manage Access" panel.

This script resolves that by calling the same REST endpoint the SharePoint UI uses:

```
POST /_api/web/Lists(@a1)/GetItemById(@a2)/ShareLink
```

with `inviteesToRemove` populated from live `GetSharingInformation` data. This removes both the SP group membership and the UI-visible invitation metadata atomically.

If the REST API call fails for any reason, the script falls back to standard SP group removal and logs a warning.

---

## Debug Mode

Set `$debug = $true` to enable verbose output:

- Every permission check attempt is logged to the console and log file.
- Group lookups, role assignments, and throttling delays are all surfaced.
- Useful for troubleshooting unexpected results or testing against a single site.

```powershell
$debug = $true
$TargetSiteUrl = "https://contoso.sharepoint.com/sites/TestSite"
$Mode = "Report"
.\SPO-UserCleanup.ps1
```

---

## Notes and Caveats

- **M365 Group membership (Teams/Groups-connected sites):** The script detects the user via M365 group membership in the Report phase, but does **not** remove them from the M365 group itself or from Teams. Use the Microsoft 365 admin center or Exchange Online PowerShell to remove a user from the M365 group if needed.
- **Site Collection Admins:** The script reports when a user is a Site Collection Admin but does not automatically remove that admin role. Manual removal via the SharePoint admin center or `Remove-PnPSiteCollectionAdmin` is recommended.
- **Entra ID group memberships:** Reported but not removed by this script. The script cannot safely remove a user from an Entra ID security group since that group may be used for purposes beyond SharePoint access.
- **OneDrive sites:** Excluded by default (`$IncludeOneDrive = $false`). Set to `$true` to include them — note this significantly increases scan time.
- **Large tenants:** For tenants with 500+ sites, a full Report scan can take several hours. Use `$TargetSiteUrl` to process individual sites or subsets.
- **Idempotency:** The Remove phase is safe to run multiple times. Attempts to remove a user who is no longer present are silently skipped.

---

## Author

**Mike Lee**  
Created: March 3, 2026
