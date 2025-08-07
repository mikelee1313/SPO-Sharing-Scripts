- **Detection mode output (CSV) cannot be used as input for remediation.**  
- Two modes: Detection (inventory/report) and Remediation (direct action).  
- Input for remediation must be a simple list of site URLs, NOT a detection report.

---

# SharePoint Online Sharing Links Management Scripts

> **⚠️ WARNING:** Removing sharing links in Remediation mode is permanent. Any previously distributed links will stop working immediately.

## Overview

These PowerShell scripts help you inventory and remediate Organization sharing links across SharePoint Online in Microsoft 365.  
- **Detection mode** inventories all sharing links and outputs a CSV report.  
- **Remediation mode** converts users with Organization sharing links to direct permissions and removes those links/groups.

There are two script versions:
- **Get-and-Remove-SPOSharingLinks-pnp2x.ps1** – For PnP.PowerShell 2.x
- **Get-and-Remove-SPOSharingLinks-pnp3x.ps1** – For PnP.PowerShell 3.x

---

## Features

- **Inventory:** Scan all or specific SharePoint sites for sharing links, focusing on Organization links.
- **Remediation:** Convert users with Organization links to direct permissions and remove those sharing groups/links.
- **Automatic Cleanup:** Removes empty/corrupted sharing groups during remediation.
- **Robust Logging:** Writes INFO, DEBUG (if enabled), and ERROR logs for audit and troubleshooting.
- **Throttling Resilience:** Built-in exponential backoff and retry to handle SharePoint/Graph throttling.
- **Flexible Input:** Remediation can target all sites or a custom list from a simple CSV file of site URLs.
- **Modern Auth:** Uses certificate-based Entra ID (Azure AD) app authentication.

---

## Prerequisites

- **PowerShell 7+**
- **PnP.PowerShell** (version 2.x or 3.x, matching the script)
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
   Get the version matching your PnP.PowerShell install.

3. **Set Up Entra ID App & Certificate**  
   - Register a new app in Entra ID (Azure AD) and upload your certificate.
   - Grant required API permissions and admin consent.

4. **Export and Note Certificate Thumbprint**  
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=SPOScripts" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable
   $cert.Thumbprint
   Export-Certificate -Cert $cert -FilePath "C:\Temp\SPOScripts.cer"
   ```

---

## Configuration

Open the script and update the variables at the top:

```powershell
$tenantname = "yourtenant"        # Without '.onmicrosoft.com'
$appID = "your-app-id"            # Entra App ID
$thumbprint = "your-cert-thumb"   # Certificate thumbprint
$tenant = "your-tenant-id"        # Directory (tenant) ID (GUID)
$searchRegion = "NAM"             # "NAM", "EUR", etc.
$Mode = "Detection"               # "Detection" or "Remediation"
$debugLogging = $false            # $true for debug logs, $false for info/errors only
$inputfile = ""                   # Optional: simple CSV of site URLs for targeted processing
```

---

## Script Parameters & Modes

| Parameter         | Type     | Required?  | Description                                                   |
|-------------------|----------|------------|---------------------------------------------------------------|
| `$tenantname`     | String   | Yes        | Your M365 tenant (no `.onmicrosoft.com`)                      |
| `$appID`          | String   | Yes        | Entra ID application ID                                       |
| `$thumbprint`     | String   | Yes        | Certificate thumbprint for authentication                     |
| `$tenant`         | String   | Yes        | Directory (tenant) ID                                         |
| `$searchRegion`   | String   | No         | Microsoft Graph search region                                 |
| `$Mode`           | String   | Yes        | `"Detection"` = report only, `"Remediation"` = convert/remove |
| `$debugLogging`   | Boolean  | No         | `$true` for debug-level logs                                  |
| `$inputfile`      | String   | No         | Optional: simple CSV of site URLs to process                  |

**Important:**
- When using `$inputfile`, it must be a plain list of site URLs (one per line, or a CSV with "URL" as the header).
- **The output CSV from Detection mode is NOT valid as input for Remediation mode.**
- Remediation mode will process ALL sites if `$inputfile` is blank.

---

## Usage

### Detection (Inventory) Mode

```powershell
$Mode = "Detection"
$inputfile = ""   # or a simple CSV list of site URLs
.\Get-and-Remove-SPOSharingLinks-pnp3x.ps1    # Or -pnp2x.ps1
```
- Outputs a CSV report of all detected sharing links.  
- **The report is for review only; it cannot be used as remediation input.**

### Remediation Mode

```powershell
$Mode = "Remediation"
$inputfile = ""   # Or a simple CSV list of SharePoint site URLs only
.\Get-and-Remove-SPOSharingLinks-pnp3x.ps1    # Or -pnp2x.ps1
```
- Converts users with Organization sharing links to direct permissions and removes links/groups for ALL specified sites.

### Custom Remediation (Subset of Sites)

If you want to remediate only certain sites:
1. Create a CSV file (e.g., `mysites.csv`) with:
   ```
   URL
   https://yourtenant.sharepoint.com/sites/site1
   https://yourtenant.sharepoint.com/sites/site2
   ```
2. Set:
   ```powershell
   $Mode = "Remediation"
   $inputfile = "C:\Path\To\mysites.csv"
   ```
3. Run the script.

---

## Output

- **CSV Report:**  
  `%TEMP%\SPO_SharingLinks_YYYY-MM-DD_HH-MM-SS.csv`  
  (Generated in Detection mode for inventory/review only.)

- **Log File:**  
  `%TEMP%\SPOSharingLinksYYYY-MM-DD_HH-MM-SS_logfile.log`  
  (INFO, DEBUG [if enabled], ERROR)

---

## Limitations & Important Notes

- **You CANNOT use the Detection mode report as input for Remediation mode.**  
  Remediation only accepts a plain list of site URLs.
- **Remediation always removes Organization sharing links after converting users.**  
  This is a permanent action!
- **Always start with Detection mode and review the report before remediation.**
- **Test in a non-production tenant before broad usage.**

---

## Troubleshooting

- **Authentication errors:** Verify app registration, certificate, and permissions.
- **Throttling:** The script automatically retries, but for large tenants, consider splitting input files.
- **Permission issues:** Confirm app permissions and admin consent.
- **Enable `$debugLogging = $true` for troubleshooting.**

---

## Security & Compliance

- Store certificates securely.
- Limit script/app access.
- Clean up temporary files after processing.
- Document script runs for compliance.

---

## FAQ

**Q: Can I use the Detection report as input for remediation?**  
A: **No.** Only a plain list of site URLs is supported for remediation.

**Q: What does Remediation do?**  
A: Converts users with Organization links to direct permissions and removes those sharing groups and links.

**Q: Can I preserve sharing links after remediation?**  
A: No. Remediation always removes Organization sharing links as part of the process.

---

## Version History

- **Aug 7, 2025:** Updated for clarity—Detection output cannot be used as remediation input. Improved throttling handling and logging.
- See script header for detailed changelog.

---

## Disclaimer

These scripts are provided **AS IS** without warranty. Test thoroughly and review all changes before running remediation.

---

**Author:** Mike Lee  
*Last updated: Aug 7, 2025*
