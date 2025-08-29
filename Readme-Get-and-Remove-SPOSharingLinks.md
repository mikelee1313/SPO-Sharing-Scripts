# SharePoint Online Sharing Links Management Script

> **⚠️ WARNING:** Removing sharing links in Remediation mode is permanent. Any previously distributed links will stop working immediately.

## Overview

This PowerShell script inventories and remediates SharePoint Online (SPO) sharing links across your Microsoft 365 tenant.  
It is designed to help organizations identify, report, and (optionally) clean up "Organization" sharing links, converting them to direct user permissions for improved security and compliance.

- **Detection mode**: Inventories all sharing links and produces a comprehensive CSV report.
- **Remediation mode**: Converts Organization sharing links to direct permissions, removes those links and groups, and auto-cleans corrupted sharing groups.

Supports scanning ALL sites or a specific subset from a CSV input file. Flexible sharing links are detected and reported but never modified.

---

## Features

- **Inventory and Report**: Scans all or selected SharePoint sites for sharing links, focusing on Organization links.
- **Remediation**: Converts Organization sharing links to direct permissions, removes sharing groups/links, and cleans up corrupted groups.
- **Flexible Input**: Processes all sites or a custom list from a CSV file (plain URLs or script-generated CSV).
- **Throttling Resilience**: Built-in retry and exponential backoff for SharePoint/Graph throttling.
- **Robust Logging**: INFO, ERROR, and optional DEBUG logs for troubleshooting.
- **Modern Authentication**: Uses certificate-based Entra ID (Azure AD) app authentication.
- **Automatic Mode Selection**: Detects if input CSV is a Detection mode output and automatically runs Remediation on Organization links only.
- **Detailed Output**: CSV report with per-link and per-site details, including search status, link removal results, and more.

---

## Prerequisites

- **PowerShell 7+**
- **PnP.PowerShell** module (v2.x or v3.x)
- **Entra ID (Azure AD) App Registration**
    - Application Permissions:
        - `SharePoint:Sites.FullControl.All`
        - `SharePoint:User.Read.All`
        - `Graph:Sites.FullControl.All`
        - `Graph:Sites.Read.All`
        - `Graph:Files.Read.All`
    - Certificate uploaded to app registration
- **SharePoint/Global Admin** permissions

---

## Installation

1. **Install PnP PowerShell**
   ```powershell
   Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
   ```

2. **Download the Script**
   - Use the latest version from this repo:  
     [`Get-and-Remove-SPOSharingLinks.ps1`](https://github.com/mikelee1313/SPO-Sharing-Scripts/blob/main/Get-and-Remove-SPOSharingLinks.ps1)

3. **Set Up Entra ID App & Certificate**
   - Register a new app in Entra ID (Azure AD), upload your certificate.
   - Grant required API permissions and admin consent.

4. **Export and Note Certificate Thumbprint**
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=SPOScripts" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable
   $cert.Thumbprint
   Export-Certificate -Cert $cert -FilePath "C:\Temp\SPOScripts.cer"
   ```

---

## Configuration

Open the script and set these variables at the top:

```powershell
$tenantname    = "yourtenant"      # Without '.onmicrosoft.com'
$appID         = "your-app-id"     # App Registration ID
$thumbprint    = "your-cert-thumb" # Certificate thumbprint
$tenant        = "your-tenant-id"  # Directory (tenant) ID (GUID)
$searchRegion  = "NAM"             # "NAM", "EUR", etc.
$Mode          = "Detection"       # "Detection" or "Remediation"
$debugLogging  = $false            # $true for debug logs, $false for info/errors only
$inputfile     = ""                # Optional: CSV file for targeted sites
```

---

## Script Parameters & Modes

| Parameter         | Type     | Required?  | Description                                                             |
|-------------------|----------|------------|-------------------------------------------------------------------------|
| `$tenantname`     | String   | Yes        | Your M365 tenant (no `.onmicrosoft.com`)                                |
| `$appID`          | String   | Yes        | Entra ID (Azure AD) application ID                                      |
| `$thumbprint`     | String   | Yes        | Certificate thumbprint for authentication                               |
| `$tenant`         | String   | Yes        | Directory (tenant) ID (GUID)                                            |
| `$searchRegion`   | String   | No         | Microsoft Graph search region                                           |
| `$Mode`           | String   | Yes        | `"Detection"` = report only, `"Remediation"` = convert/remove           |
| `$debugLogging`   | Boolean  | No         | `$true` for debug-level logs                                            |
| `$inputfile`      | String   | No         | Optional: CSV file (list of site URLs or Detection mode output)          |

**Important:**
- `$inputfile` can be:
  - Simple list of site URLs (one per line or a CSV with "URL" header)
  - Output CSV from a previous Detection run (automatically switches to Remediation mode for Organization links only)
- If not specified, all active sites in the tenant are processed.

---

## Usage

### Detection (Inventory) Mode

```powershell
$Mode = "Detection"
$inputfile = ""   # or a simple CSV list of site URLs
.\Get-and-Remove-SPOSharingLinks.ps1
```
- Outputs a CSV report of all detected sharing links.
- The report includes details like site owner, group, file, sharing type, link expiration, etc.

### Remediation Mode

```powershell
$Mode = "Remediation"
$inputfile = ""   # Or a simple CSV list of site URLs
.\Get-and-Remove-SPOSharingLinks.ps1
```
- Converts users with Organization sharing links to direct permissions and removes the links/groups.

### Targeted Remediation (Subset of Sites)

- Create a CSV file (e.g., `mysites.csv`) with:
  ```
  URL
  https://yourtenant.sharepoint.com/sites/site1
  https://yourtenant.sharepoint.com/sites/site2
  ```
- Set:
  ```powershell
  $Mode = "Remediation"
  $inputfile = "C:\Path\To\mysites.csv"
  ```
- Run the script.

### Remediation from Detection Output

- Run Detection mode to produce a CSV.
- Use the output CSV as `$inputfile`; the script auto-enables Remediation and processes only Organization links.

---

## Output

- **CSV Report:**  
  `%TEMP%\SPO_SharingLinks_YYYY-MM-DD_HH-MM-SS.csv`
    - Columns: Site URL, Site Owner, IB Mode, IB Segment, Template, Sharing Group Name, Members, File URL, File Owner, Filename, SharingType, Sharing Link URL, Link Expiration Date, IsTeamsConnected, SharingCapability, Last Content Modified, Search Status, Link Removed

- **Log File:**  
  `%TEMP%\SPOSharingLinksYYYY-MM-DD_HH-MM-SS_logfile.log`
    - INFO, DEBUG (if enabled), ERROR

---

## Limitations & Notes

- **Remediation is permanent:** Organization sharing links are converted and removed.
- **Detection mode output CSV can be used as input for remediation** (Organization links only)—the script detects and auto-switches mode.
- **Flexible sharing links are never modified in remediation.**
- **Always run Detection mode and review the report before remediation.**
- **Test in a non-production tenant before broad usage.**

---

## Troubleshooting

- **Authentication errors:** Check app registration, certificate, permissions.
- **Throttling:** The script auto-retries, but split large tenants into smaller batches if needed.
- **Permission issues:** Confirm API permissions and admin consent.
- **Enable `$debugLogging = $true` for detailed logs.**

---

## Security & Compliance

- Store certificates securely.
- Limit script/app access to administrators.
- Clean up temporary files after processing.
- Document all script runs for compliance records.

---

## FAQ

**Q: Can I use the Detection report as input for remediation?**  
A: Yes! The script detects its own output and runs focused remediation on Organization links only.

**Q: What does Remediation do?**  
A: Converts Organization sharing links to direct permissions and removes associated sharing groups and links.

**Q: Can I preserve sharing links after remediation?**  
A: No. Remediation always removes Organization sharing links after conversion.

**Q: Does the script modify Flexible sharing links?**  
A: No, Flexible links are only reported, never modified.

---

## Version History

- **Aug 29, 2025:** Major update—Detection output now accepted as remediation input; improved throttling handling, logging, and input flexibility.
- See script header for detailed changelog.

---

## Disclaimer

These scripts are provided **AS IS** without warranty. Test thoroughly and review all changes before running remediation.

---

---

Let me know if you'd like this pushed as a file to your repo or need further customization!
