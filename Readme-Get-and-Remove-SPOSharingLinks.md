# SharePoint Online Sharing Links Management Script

## Overview

The **Get-and-Remove-SPOSharingLinks-pnpxx.ps1** script is a comprehensive PowerShell tool designed to identify, inventory, and remediate SharePoint Online sharing links across your Microsoft 365 tenant. It focuses specifically on **Organization sharing links** and provides a two-step workflow for safe and efficient remediation.

### Key Features

- 🔍 **Complete Inventory**: Scans all SharePoint sites to identify sharing links
- 🎯 **Targeted Remediation**: Converts Organization sharing links to direct permissions
- 🔄 **Flexible Link Management**: Option to preserve or remove sharing links after user conversion
- 🧹 **Automatic Cleanup**: Removes corrupted and empty sharing groups
- 📊 **Detailed Reporting**: Generates comprehensive CSV reports
- 🚀 **Two-Step Workflow**: Report first, then remediate based on findings
- 🔄 **Smart Detection**: Automatically recognizes its own output for targeted processing

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Installation](#installation)
3. [Configuration](#configuration)
4. [Usage Workflows](#usage-workflows)
5. [Script Parameters](#script-parameters)
6. [Output Files](#output-files)
7. [Common Scenarios](#common-scenarios)
8. [Troubleshooting](#troubleshooting)
9. [Security Considerations](#security-considerations)
10. [Support](#support)

---

## Prerequisites

### Required Software
- **PowerShell 7+**
- **PnP.PowerShell module - Note: Use Compatible Version of Script**
- **Get-and-Remove-SPOSharingLinks-pnp2x.ps1 (PNP2.x)**
- **Get-and-Remove-SPOSharingLinks-pnp3x.ps1 (PNP3.x)**

### Microsoft 365 Requirements
- **SharePoint Online** subscription
- **Global Administrator** or **SharePoint Administrator** permissions
- **Entra ID Application** with certificate-based authentication

### Permissions Required
The Entra ID application must have the following **Application permissions**:
- `SharePoint:Sites.FullControl.All` - Access to all SharePoint sites
- `SharePoint:User.Read.All` - Read user profiles and group memberships
- `Graph:Sites.FullControl.All` - Access to all SharePoint sites via Graph
- `Graph:Files.Read.All` - Access to all SharePoint files via Graph

---

## Installation

### Step 1: Install PnP PowerShell Module

```powershell
# Install the latest PnP PowerShell module
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force

# Verify installation
Get-Module -Name PnP.PowerShell -ListAvailable
```

### Step 2: Download the Script

1. Download `Get-and-Remove-SPOSharingLinks-pnpxx.ps1` to your local machine
2. Save it in a dedicated folder (e.g., `C:\SharePointScripts\`)

### Step 3: Create Entra ID Application

1. **Navigate to Azure Portal** → **Entra ID** → **App registrations**
2. **Click "New registration"**
3. **Configure the application**:
   - Name: `SharePoint Sharing Links Management`
   - Supported account types: `Accounts in this organizational directory only`
   - Redirect URI: Leave blank
4. **Note the Application (client) ID** and **Directory (tenant) ID**

### Step 4: Configure API Permissions

1. **Go to API permissions** → **Add a permission**
2. **Select Microsoft Graph** → **Application permissions**
3. **Add these permissions**:
- `SharePoint:Sites.FullControl.All`
- `SharePoint:User.Read.All`
- `Graph:Sites.FullControl.All`
- `Graph:Files.Read.All`
4. **Click "Grant admin consent"**

### Step 5: Create Certificate

```powershell
# Create a self-signed certificate for authentication
$cert = New-SelfSignedCertificate -Subject "CN=SharePointSharingLinksApp" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256

# Note the thumbprint for script configuration
$cert.Thumbprint
```

### Step 6: Upload Certificate to Entra ID

1. **In your app registration** → **Certificates & secrets**
2. **Click "Upload certificate"**
3. **Export the certificate** (without private key) and upload the `.cer` file
4. **Note the certificate thumbprint**

---

## Configuration

### Edit Script Variables

Open `Get-and-Remove-SPOSharingLinks-pnpxx.ps1` and update these variables:

```powershell
# ----------------------------------------------
# Set Variables - EDIT THESE VALUES
# ----------------------------------------------
$tenantname = "yourcompany"                                     # Your tenant name (without .onmicrosoft.com)
$appID = "your-app-id-here"                                     # Your Entra App ID
$thumbprint = "your-certificate-thumbprint-here"               # Your certificate thumbprint
$tenant = "your-tenant-id-here"                                # Your Tenant ID (GUID)
$searchRegion = "NAM"                                          # Your region: NAM, EUR, or APAC
```

### Key Configuration Options

| Parameter | Description | Default | Recommendations |
|-----------|-------------|---------|-----------------|
| `$convertOrganizationLinks` | Enable remediation mode | `$false` | Start with `$false` for reporting |
| `$debugLogging` | Enable detailed logging | `$true` | Keep `$true` for initial runs |
| `$inputfile` | Path to input CSV file | `$null` | Comment out for full tenant scan |

---

## Usage Workflows

### 🔄 Recommended Two-Step Workflow

This is the **safest and most efficient** approach:

#### **Step 1: Generate Report (Inventory Mode)**

```powershell
# 1. Edit the script configuration
$convertOrganizationLinks = $false          # Report mode
$debugLogging = $true                       # Enable detailed logging
$inputfile = $null                          # Scan all sites

Note: To run Inventory Mode against a list of sites, populate the $inputfile with your own site list.

# 2. Run the script
.\Get-and-Remove-SPOSharingLinks-pnpxx.ps1

# 3. Review the generated CSV file in %TEMP%
# File name format: SPO_SharingLinks_YYYY-MM-DD_HH-MM-SS.csv
```

#### **Step 2: Remediate Using Report (Automatic Mode)**

```powershell
# 1. Edit the script configuration
$inputfile = "C:\Temp\SPO_SharingLinks_2025-07-01_14-30-15.csv"

# 2. Run the script (it will auto-enable remediation for Organization links)
.\Get-and-Remove-SPOSharingLinks-pnpxx.ps1

# Note: The script automatically:
# - Detects its own CSV format
# - Filters for Organization sharing links only
# - Enables remediation mode ($convertOrganizationLinks = $true)
# - By default, removes sharing links ($RemoveSharingLink = $true)
# - Enables cleanup mode ($cleanupCorruptedSharingGroups = $true)
```

#### **Step 3: Remediate While Preserving Sharing Links (Optional)**

```powershell
# If you want to convert users to direct permissions but KEEP the sharing links:
$inputfile = "C:\Temp\SPO_SharingLinks_2025-07-01_14-30-15.csv"
$RemoveSharingLink = $false  # This preserves the sharing links

# Run the script
.\Get-and-Remove-SPOSharingLinks-pnpxx.ps1

# Note: With $RemoveSharingLink = $false:
# - Users are still converted to direct permissions
# - Sharing links remain intact and functional
# - Corrupted sharing groups are NOT cleaned up
```

### 📊 Alternative Workflows

#### **Direct Remediation with Site List**

```powershell
# Create a simple CSV with site URLs
# File: sitelist.csv
# Content:
# URL
# https://yourcompany.sharepoint.com/sites/site1
# https://yourcompany.sharepoint.com/sites/site2

$inputfile = "C:\temp\sitelist.csv"
$convertOrganizationLinks = $true
.\Get-and-Remove-SPOSharingLinks-pnp2x.ps1
```

#### **Full Tenant Scan with Immediate Remediation**

```powershell
# ⚠️ WARNING: This processes ALL sites immediately
$convertOrganizationLinks = $true
$inputfile = $null
.\Get-and-Remove-SPOSharingLinks-pnpxx.ps1
```

---

## Script Parameters

### Core Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `$tenantname` | String | ✅ Yes | Your M365 tenant name (without .onmicrosoft.com) |
| `$appID` | String | ✅ Yes | Entra ID Application ID |
| `$thumbprint` | String | ✅ Yes | Certificate thumbprint for authentication |
| `$tenant` | String | ✅ Yes | Tenant ID (GUID) |

### Operational Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `$convertOrganizationLinks` | Boolean | `$false` | Enable remediation mode |
| `$RemoveSharingLink` | Boolean | `$true` | When `$true`, removes sharing links after converting users. When `$false`, preserves sharing links while still converting users to direct permissions. |
| `$cleanupCorruptedSharingGroups` | Boolean | `$false`* | Clean up empty sharing groups |
| `$debugLogging` | Boolean | `$true` | Enable detailed logging |
| `$inputfile` | String | `$null` | Path to input CSV file |
| `$searchRegion` | String | `"NAM"` | Microsoft Graph search region |

*Automatically set to `$true` when `$convertOrganizationLinks` is `$true` AND `$RemoveSharingLink` is `$true`

### Region Codes

| Region | Code | Description |
|--------|------|-------------|
| North America | `NAM` | United States, Canada |
| Europe | `EUR` | European Union countries |
| Asia Pacific | `APAC` | Asia and Pacific regions |

---

## Output Files

### CSV Report File

**Location**: `%TEMP%\SPO_SharingLinks_YYYY-MM-DD_HH-MM-SS.csv`

**Columns**:
- `Site URL` - SharePoint site URL
- `Site Owner` - Site collection owner
- `IB Mode` - Information Barrier mode
- `IB Segment` - Information Barrier segments
- `Site Template` - SharePoint template type
- `Sharing Group Name` - Name of the sharing group
- `Sharing Link Members` - Users with access (Name <Email>)
- `File URL` - Direct link to shared document
- `File Owner` - Document owner
- `Sharing Link URL` - Current Sharing Link (to be removed)
- `IsTeamsConnected` - Whether site is Teams-connected
- `SharingCapability` - Site sharing settings
- `Last Content Modified` - When content was last modified
- `Link Removed` - Whether sharing link was removed (True/False)

### Log File

**Location**: `%TEMP%\SPOSharingLinksYYYY-MM-DD_HH-MM-SS_logfile.log`

**Log Levels**:
- `[INFO]` - General operational information
- `[DEBUG]` - Detailed technical information (when enabled)
- `[ERROR]` - Errors and warnings

---

## Common Scenarios

### 🎯 Scenario 1: Monthly Sharing Links Audit

```powershell
# Run monthly report to track sharing links
$convertOrganizationLinks = $false
$debugLogging = $false                    # Reduce log verbosity for regular runs
.\Get-and-Remove-SPOSharingLinks-pnpxx.ps1
```

### 🔧 Scenario 2: Remediate Specific Sites

```powershell
# Create CSV with problematic sites
# Then run remediation
$inputfile = "C:\temp\problematic_sites.csv"
$convertOrganizationLinks = $true
$RemoveSharingLink = $true  # Default: removes sharing links after user conversion
.\Get-and-Remove-SPOSharingLinks-pnpxx.ps1
```

### � Scenario 3: Convert Users but Preserve Sharing Links

```powershell
# Convert Organization sharing link users to direct permissions
# But KEEP the sharing links intact
$convertOrganizationLinks = $true
$RemoveSharingLink = $false  # Preserves sharing links
$inputfile = "C:\temp\sites_to_process.csv"
.\Get-and-Remove-SPOSharingLinks-pnpxx.ps1
```

### �📊 Scenario 4: Executive Dashboard Data

```powershell
# Generate data for executive reporting
$convertOrganizationLinks = $false
$debugLogging = $false
# Process the CSV with Power BI or Excel for visualization
```

### 🧹 Scenario 4: Cleanup After Migration

```powershell
# After migrating from external sharing to direct permissions
# Use the script's CSV output to verify Organization links are converted
$inputfile = "previous_scan_results.csv"
# Script auto-detects and processes Organization links only
```

---

## Troubleshooting

### Common Issues

#### ❌ Authentication Errors

**Error**: `AADSTS70011: The provided value for the input parameter 'scope' is not valid`

**Solution**:
1. Verify the app has correct permissions
2. Ensure admin consent is granted
3. Check certificate is properly uploaded

#### ❌ Certificate Issues

**Error**: `Certificate with thumbprint 'xxx' not found`

**Solutions**:
```powershell
# List available certificates
Get-ChildItem -Path "Cert:\CurrentUser\My"

# Check if certificate is in correct store
Get-ChildItem -Path "Cert:\LocalMachine\My"
```

#### ❌ Throttling Errors

**Error**: `Too many requests` or `Request limit exceeded`

**Solution**: The script automatically handles throttling with exponential backoff. For severe throttling:
1. Reduce batch sizes by using input files with fewer sites
2. Run during off-peak hours
3. Increase delays in the throttling function

#### ❌ Permission Errors

**Error**: `Access denied` when processing sites

**Solutions**:
1. Verify app has `Sites.FullControl.All` permission
2. Check that admin consent is properly granted
3. Ensure the app is not blocked by conditional access policies

### Performance Optimization

#### For Large Tenants (1000+ sites)

1. **Use input files** to process sites in batches
2. **Run during off-peak hours** to minimize throttling
3. **Disable debug logging** for production runs
4. **Monitor throttling** and adjust timing if needed

#### Memory Management

```powershell
# For very large tenants, restart PowerShell session periodically
# The script processes sites one at a time to minimize memory usage
```

---

## Security Considerations

### 🔒 Certificate Security

- **Store certificates securely** in the Windows Certificate Store
- **Use strong passwords** for certificate export if needed
- **Rotate certificates regularly** (annually recommended)
- **Limit certificate access** to authorized administrators only

### 🛡️ Application Security

- **Review app permissions regularly**
- **Monitor app usage** through Entra ID audit logs
- **Use descriptive app names** for easy identification
- **Document app ownership** and purpose

### 📋 Data Protection

- **Secure log files** - they contain user and site information
- **Clean up temporary files** after script execution
- **Follow data retention policies** for generated reports
- **Encrypt sensitive data** if storing long-term

### 🔍 Audit Recommendations

1. **Log all script executions** with dates and operators
2. **Review remediation results** before final approval
3. **Test in non-production environment** first
4. **Maintain change documentation** for compliance

---

## Script Workflow Diagram

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   Start Script  │ -> │  Load Variables │ -> │   Authenticate  │
└─────────────────┘    └─────────────────┘    └─────────────────┘
                                                       │
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   Get Site List │ <- │  Check Input    │ <- │  Connect to     │
│   (All Sites)   │    │  File Type      │    │  Admin Center   │
└─────────────────┘    └─────────────────┘    └─────────────────┘
         │                       │
         │              ┌─────────────────┐
         │              │   Load CSV      │
         │              │   Filter Org    │
         │              │   Links Only    │
         │              └─────────────────┘
         │                       │
         ↓                       ↓
┌─────────────────────────────────────────────────────────────────┐
│                    Process Each Site                            │
│  1. Connect to site                                             │
│  2. Get SharePoint groups                                       │
│  3. Get group members                                           │
│  4. If remediation mode:                                        │
│     - Remove users from sharing groups                         │
│     - Grant direct permissions                                  │
│     - If $RemoveSharingLink = $true:                            │
│       * Remove sharing links                                   │
│       * Clean up empty sharing groups                          │
│  5. Write data to CSV                                           │
└─────────────────────────────────────────────────────────────────┘
                                │
                       ┌─────────────────┐
                       │  Generate Final │
                       │     Report      │
                       │   (CSV + Log)   │
                       └─────────────────┘
```

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 2.1 | July 2, 2025 | Added `$RemoveSharingLink` parameter to allow preserving sharing links while converting users |
| 2.0 | July 1, 2025 | Complete rewrite with two-step workflow |
| 1.5 | June 2025 | Added automatic cleanup integration |
| 1.0 | May 2025 | Initial release |

---

## Support

### Getting Help

1. **Check the log files** for detailed error information
2. **Review this README** for common scenarios
3. **Test in a small environment** before large-scale deployment
4. **Document your specific use case** when seeking support

### Best Practices

- ✅ **Always test in non-production first**
- ✅ **Start with report mode** before remediation
- ✅ **Review CSV output** before running remediation
- ✅ **Consider whether to preserve sharing links** by setting `$RemoveSharingLink = $false`
- ✅ **Backup important data** before making changes
- ✅ **Run during maintenance windows** for large operations
- ✅ **Monitor performance** and adjust timing as needed

### Script Maintenance

- 📅 **Review quarterly** for SharePoint API changes
- 🔄 **Update PnP PowerShell** module regularly
- 📊 **Monitor execution metrics** for performance trends
- 🔒 **Rotate certificates** annually
- 📝 **Update documentation** with lessons learned

---

## Disclaimer

This sample script is provided **AS IS** without warranty of any kind. Microsoft and the script authors disclaim all implied warranties including, without limitation, any implied warranties of merchantability or fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you.

---

*Last Updated: July 2, 2025*  
*Script Version: 2.1*  
*Author: Mike Lee*
