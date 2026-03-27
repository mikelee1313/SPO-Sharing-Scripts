# SPO-permissionsAudit.ps1

Audits SharePoint Online permission changes across a list of site collections using the Microsoft 365 Unified Audit Log (`Search-UnifiedAuditLog`).

**Authors:** Mike Lee / Mariel Williams  
**Version:** 1.0 | 3/27/26

---

## Overview

This script queries the Unified Audit Log for all permission-related events across one or more SharePoint Online site collections and exports a clean, admin-readable CSV report. It is designed for SharePoint admins who need to audit who changed permissions, on what, and when — without having to parse raw JSON from the Purview compliance portal.

---

## Prerequisites

| Requirement | Details |
|---|---|
| PowerShell | 5.1 or 7+ |
| Module | `ExchangeOnlineManagement` (`Install-Module ExchangeOnlineManagement`) |
| Role | Must have **Audit Logs** or **View-Only Audit Logs** role in Microsoft 365 |
| Licensing | Microsoft 365 tenant with Unified Audit Log enabled |

The script will automatically call `Connect-ExchangeOnline` if a session is not already active.

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-SiteListPath` | Yes | `C:\temp\SPOSiteList.txt` | Path to a text file with one SPO site URL per line |
| `-StartDate` | No | 3 days ago (UTC) | Start of the audit window |
| `-EndDate` | No | Now (UTC) | End of the audit window |
| `-OutputPath` | No | `.\SPO-PermissionsAudit_<timestamp>.csv` | Path for the exported CSV |
| `-ResultSize` | No | `5000` | Records per page (max 5000) |
| `-IncludeSystemEvents` | No | Off | Include internal SPO system-group side-effect rows |

---

## Site List File Format

Create a plain text file with one SharePoint site URL per line. Blank lines and lines without a URL are ignored.

```text
https://contoso.sharepoint.com/sites/Finance
https://contoso.sharepoint.com/sites/HR
https://contoso.sharepoint.com/teams/ProjectAlpha
```

---

## Usage

```powershell
# Basic – last 3 days, default site list path
.\SPO-permissionsAudit.ps1

# Custom date range
.\SPO-permissionsAudit.ps1 -SiteListPath .\sites.txt -StartDate (Get-Date).AddDays(-30)

# Full custom run
.\SPO-permissionsAudit.ps1 `
    -SiteListPath  "C:\Audit\sites.txt" `
    -StartDate     (Get-Date).AddDays(-90) `
    -EndDate       (Get-Date) `
    -OutputPath    "C:\Audit\Results.csv"

# Include internal SPO system-group events
.\SPO-permissionsAudit.ps1 -SiteListPath .\sites.txt -IncludeSystemEvents
```

---

## Events Captured (19 operation types)

| Category | Operation | Friendly Label |
|---|---|---|
| **Unique permissions** | `SharingInheritanceBroken` | Unique Permissions Created (Inheritance Broken) |
| | `UniquePermissionsSet` | Unique Permissions Set |
| **Group membership** | `AddedToGroup` | User Added to Group |
| | `RemovedFromGroup` | User Removed from Group |
| **Admin changes** | `SiteCollectionAdminAdded` | Site Collection Admin Added |
| | `SiteCollectionAdminRemoved` | Site Collection Admin Removed |
| **Permission levels** | `PermissionLevelAdded` | Permission Level Added |
| | `PermissionLevelChanged` | Permission Level Modified |
| **Direct sharing** | `SharingSet` | Permissions Granted |
| | `SharingPermissionChanged` | Permission Changed |
| **Sharing links** | `SharingLinkCreated` | Sharing Link Created |
| | `AddedToSharingLink` | User Added to Sharing Link |
| **Specific-people links** | `SecureLinkCreated` | Secure Link Created (Specific People) |
| | `SecureLinkUpdated` | Secure Link Updated (Specific People) |
| | `AddedToSecureLink` | User Added to Secure Link |
| | `RemovedFromSecureLink` | User Removed from Secure Link |
| **Anonymous links** | `AnonymousLinkCreated` | Anonymous Link Created |
| | `AnonymousLinkUpdated` | Anonymous Link Updated |
| | `AnonymousLinkRemoved` | Anonymous Link Removed |

---

## Output CSV Columns

| Column | Description |
|---|---|
| `DateTime` | When the event occurred (UTC) |
| `PerformedBy` | UPN of the user who made the change |
| `Action` | Human-readable event description |
| `ItemType` | File, Folder, Web, List |
| `RelativePath` | Relative path of the affected item (e.g. `Shared Documents/Folder1`) |
| `SiteUrl` | Site collection URL |
| `AffectedUser` | User or group that access was granted/removed for |
| `PermissionsGranted` | Permission level granted (e.g. `Contribute`, `Edit`, `Limited Access`) |
| `GroupAffected` | SharePoint group name involved in the change |
| `LinkScope` | `SpecificPeople`, `Anonymous`, etc. (blank if not a link event) |
| `ClientIP` | IP address of the actor |

### System Event Filtering

SPO automatically generates several internal `AddedToGroup` events (e.g. `Limited Access System Group`) as side effects of sharing operations. These are **filtered out by default** to reduce noise.

Use `-IncludeSystemEvents` to include them. The console output will always show how many were suppressed.

---

## Console Output Example

```
SPO Permissions Audit
  Sites    : 3
  Window   : 2026-02-25 00:00:00Z  ->  2026-03-27 00:00:00Z
  Operations: 19 event types

  [+] https://contoso.sharepoint.com/sites/Finance  –  14 event(s)
  [+] https://contoso.sharepoint.com/sites/HR       –  6 event(s)
  [ ] https://contoso.sharepoint.com/teams/ProjectX –  no events found

Results  : 20 permission change events
Filtered : 12 internal SPO system-group events suppressed (use -IncludeSystemEvents to include)
Exported : .\SPO-PermissionsAudit_20260327_120000.csv

── Events by action ──────────────────────────────────────────────
Action                                             Count
------                                             -----
Unique Permissions Created (Inheritance Broken)        8
Permissions Granted                                    6
User Added to Secure Link                              4
Secure Link Created (Specific People)                  2

── Events per site ───────────────────────────────────────────────
── Changes by user ───────────────────────────────────────────────
```

---

## Notes

- The script uses `SessionCommand ReturnLargeSet` with automatic paging, so sites with more than 5,000 matching events will still return all records.
- The `-ObjectIds` filter scopes each query precisely to the site URL, preventing cross-site result bleed.
- Unified Audit Log data is typically available within 30 minutes of an event but can take up to 24 hours in some cases.
- The audit log retention window depends on your Microsoft 365 licensing (90 days for E3, up to 1 year / 10 years with add-ons).

---

## License

MIT
