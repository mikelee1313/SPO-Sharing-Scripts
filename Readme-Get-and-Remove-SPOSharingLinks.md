# SharePoint Online Sharing Links Management Scripts

> **⚠️ WARNING:** Removing sharing links (`Remediation` mode) is a permanent action. Any links previously shared will stop working immediately.

## Overview

These PowerShell scripts provide a robust, auditable way to inventory and remediate sharing links across all SharePoint Online sites in your Microsoft 365 tenant. They support a two-step workflow—first reporting, then selectively converting Organization sharing links to direct permissions and optionally removing sharing links and cleaning up empty/corrupted sharing groups.

- **Get-and-Remove-SPOSharingLinks-pnp2x.ps1** – For PnP.PowerShell 2.x
- **Get-and-Remove-SPOSharingLinks-pnp3x.ps1** – For PnP.PowerShell 3.x

---

## Features

- **Comprehensive Inventory:** Scans all or selected SharePoint sites for sharing links, focusing on Organization links.
- **Targeted Remediation:** Converts users with Organization links into direct document or site permissions.
- **Safe Two-Step Workflow:** Run in Detection (report-only) mode first, then Remediation mode based on your review of the results.
- **Automatic Group Cleanup:** Removes empty or corrupted sharing groups in remediation.
- **Throttling-Resilient:** Built-in retry and exponential backoff logic for large environments.
- **Flexible Input:** Process all sites, a list of URLs, or a previous report CSV for precise targeting.
- **Detailed Output & Logging:** CSV report and log file (INFO, DEBUG, ERROR), suitable for audits.
- **Supports Modern Authentication:** Uses certificate-based Entra ID (Azure AD) app authentication.

---

## Prerequisites

- **PowerShell 7+**
- **PnP.PowerShell** (version 2.x or 3.x, matching your chosen script)
- **Microsoft 365 Tenant** with SharePoint Online
- **Entra ID Application Registration** (with certificate)
    - Application Permissions:
        - `SharePoint:Sites.FullControl.All`
        - `SharePoint:User.Read.All`
        - `Graph:Sites.FullControl.All`
        - `Graph:Sites.Read.All`
        - `Graph:Files.Read.All`
    - Certificate uploaded to the app registration

- **Administrator Permissions:** You must be a SharePoint or Global Administrator.

---

## Installation

1. **Install PnP PowerShell**  
   ```powershell
   Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
   ```

2. **Download the Script**  
   Download the script version matching your PnP.PowerShell installation.

3. **Set Up Entra ID App & Certificate**  
   - Register a new app in Entra ID (Azure AD); upload your certificate.
   - Grant required API permissions and admin consent.

4. **Export and Note Certificate Thumbprint**  
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=SPOScripts" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable
   $cert.Thumbprint
   # Export public part to .cer for upload
   Export-Certificate -Cert $cert -FilePath "C:\Temp\SPOScripts.cer"
   ```

---

## Configuration

Open the script in your editor. Update the following variables at the top:

```powershell
$tenantname = "yourtenant"        # Without '.onmicrosoft.com'
$appID = "your-app-id"            # Entra App ID
$thumbprint = "your-cert-thumb"   # Certificate thumbprint
$tenant = "your-tenant-id"        # Directory (tenant) ID (GUID)
$searchRegion = "NAM"             # "NAM", "EUR", etc.
$Mode = "Detection"               # "Detection" (report) or "Remediation" (convert/remove)
$debugLogging = $false            # $true for verbose logs, $false for info/errors only
$inputfile = ""                   # Optional path to CSV of site URLs or previous report
```

---

## Script Parameters & Modes

| Parameter                    | Type     | Default     | Description                                                                      |
|------------------------------|----------|-------------|----------------------------------------------------------------------------------|
| `$tenantname`                | String   | (required)  | Your M365 tenant (no `.onmicrosoft.com`)                                         |
| `$appID`                     | String   | (required)  | Entra ID application ID                                                          |
| `$thumbprint`                | String   | (required)  | Certificate thumbprint for authentication                                        |
| `$tenant`                    | String   | (required)  | Directory (tenant) ID                                                            |
| `$searchRegion`              | String   | "NAM"       | Microsoft Graph search region                                                    |
| `$Mode`                      | String   | "Detection" | "Detection" = report only, "Remediation" = convert/remove org links              |
| `$debugLogging`              | Boolean  | $false      | $true for debug-level logs                                                       |
| `$inputfile`                 | String   | ""          | Optional path to CSV of site URLs or script CSV output for targeted remediation   |

**Automatic Behavior:**
- If you provide a previous report CSV as `$inputfile`, the script will:
    - Filter for sites with Organization sharing links only
    - Automatically switch to Remediation mode

---

## Usage

### Step 1: Inventory (Detection Mode)

```powershell
# Set as follows:
$Mode = "Detection"
$inputfile = ""   # or provide a list of site URLs (CSV, one per line or with "URL" header)
.\Get-and-Remove-SPOSharingLinks-pnp3x.ps1    # or -pnp2x.ps1
```
- Generates a detailed CSV in your `%TEMP%` folder listing all discovered sharing links.

### Step 2: Review Report

- Open the CSV report.
    - "Sharing Group Name" containing "Organization" are candidates for remediation.
    - Review links, users, and impacted files.

### Step 3: Remediate (Convert/Remove Organization Links)

```powershell
# Use the previous report for targeted remediation:
$Mode = "Detection"           # The script will auto-enable Remediation mode for its own CSV format
$inputfile = "C:\Temp\SPO_SharingLinks_YYYY-MM-DD_HH-MM-SS.csv"
.\Get-and-Remove-SPOSharingLinks-pnp3x.ps1
```
- Only Organization links are targeted. The script converts users to direct permissions and removes the corresponding sharing groups and links.

### Step 4: Optional – Full Tenant Remediation

```powershell
$Mode = "Remediation"
$inputfile = ""   # Leave empty to process all sites
.\Get-and-Remove-SPOSharingLinks-pnp3x.ps1
```
- **CAUTION:** This will process ALL sites in the tenant and remove all Organization sharing links.

---

## Output

- **CSV Report:**  
  `%TEMP%\SPO_SharingLinks_YYYY-MM-DD_HH-MM-SS.csv`  
  Columns: Site URL, Site Owner, Sharing Group Name, Members, File/Item URL, Owner, Link URL, Link Expiration, and more.

- **Log File:**  
  `%TEMP%\SPOSharingLinksYYYY-MM-DD_HH-MM-SS_logfile.log`  
  Levels: INFO, DEBUG (if enabled), ERROR

---

## Best Practices

- **Always start in Detection mode.** Review the report before any remediation.
- **Use input files** to target specific sites or Organization links.
- **Test in a non-production tenant** before broad use.
- **Keep log files** for audit and troubleshooting.
- **Preserve sharing links** if unsure: Only set `$Mode = "Remediation"` once you are confident.

---

## Troubleshooting

- **Authentication errors:** Ensure your app registration, certificate, and permissions are correct.
- **Throttling:** The script automatically retries, but for large tenants, consider splitting input files.
- **Permission issues:** Confirm app permissions and admin consent.
- **Debug logs:** Set `$debugLogging = $true` for troubleshooting.

---

## Security & Compliance

- Store certificates securely.
- Limit access to the app registration and script.
- Clean up temporary files after processing.
- Document changes and script runs for compliance.

---

## FAQ

**Q: Can I convert users but keep sharing links active?**  
A: No. Remediation mode always removes Organization sharing links after converting users.

**Q: Will the script process all sites?**  
A: Yes, unless you specify an input CSV to limit the scope.

**Q: What does "Organization" link mean?**  
A: Sharing links accessible by anyone in your organization (tenant-wide).

---

## Version History

- **August 7, 2025:** Major update—automatic remediation mode for CSV inputs, robust throttling handling, and improved reporting.
- See script header for authorship and detailed changelog.

---

## Disclaimer

These scripts are provided **AS IS** without warranty. Use at your own risk. Test in non-production and review all changes before running remediation.

---

**Author:** Mike Lee  
*Last updated: August 7, 2025*
