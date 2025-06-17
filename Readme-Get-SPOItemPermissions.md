

---

# Get-SPOItemPermissions.ps1

## Overview

`Get-SPOItemPermissions.ps1` is a comprehensive PowerShell script for scanning SharePoint Online (SPO) sites to enumerate all files and folders, collecting detailed permissions information—who has access, their roles, and whether permissions are inherited or unique. Results are exported to an Excel file for easy auditing and review.

**Author:** Mike Lee  
**Created:** 6/17/25

---

## Features

- Connects to multiple SharePoint Online sites using app-only authentication.
- Recursively scans all document libraries and lists (excluding system/irrelevant folders).
- Collects detailed permission info for every file and folder:
  - Users and groups with access
  - Roles assigned
  - Inherited vs. unique permissions
  - Creator and creation date
- Outputs:
  - Excel file with all item permissions
  - Log file documenting script activity
  - Optional summary worksheet

---

## Prerequisites

1. **PowerShell Modules** (installed automatically if missing):
   - [ImportExcel](https://github.com/dfinke/ImportExcel)
   - [PnP.PowerShell](https://github.com/pnp/powershell)
2. **App Registration** in Entra ID (Azure AD):
   - Permissions: `Sites.FullControl.All` recommended
   - Certificate-based authentication (App ID, Thumbprint, Tenant ID)
3. **Input File:**
   - Plain text file (`C:\temp\SPOSiteList.txt` by default) containing SPO site URLs (one per line)
4. **PowerShell 7+** recommended

---

## Configuration

Edit the top of the script to provide:

```powershell
# --- Tenant and App Registration Details ---
$appID = "your-app-id"                  # Entra App ID
$thumbprint = "your-cert-thumbprint"    # Certificate thumbprint
$tenant = "your-tenant-id"              # Tenant ID

# --- Input File Path ---
$inputFilePath = 'C:\temp\SPOSiteList.txt'  # Path to site URLs file

# --- Script Behavior Settings (optional) ---
$batchSize = 100
$maxItemsPerSheet = 5000
```

---

## Usage

1. **Edit configuration variables** near the top of the script.
2. **Prepare the input file:** List SPO site URLs, one per line.
3. **Run the script:**

   ```powershell
   .\Get-SPOItemPermissions.ps1
   ```

   (Run from an elevated PowerShell prompt if needed.)

4. **Results:**
   - Excel file: `%TEMP%\All_Item_Permissions_[timestamp].xlsx`
   - Log file: `%TEMP%\All_Item_Permissions_[timestamp].txt`
   - Excel file will open automatically if possible.

---

## Output Details

- **Excel File:** Each worksheet contains a batch of item permissions with columns for Site URL, Item Type, Library Name, Path, Name, Created By, Date, Permission Type (Unique/Inherited), Users, and Roles.
- **Summary:** Optionally, a summary worksheet lists total items processed, number of sites, unique permissions, and duration.
- **Log File:** Tracks all actions, errors, and warnings for troubleshooting.

 ![image](https://github.com/user-attachments/assets/9f6b0a64-bb36-45d0-b384-20acbefc9d4a)

 ![image](https://github.com/user-attachments/assets/f7be3d95-cef8-4dbc-a3f7-df0e5331425e)

 ![image](https://github.com/user-attachments/assets/d87b7cdb-7092-4358-8fb3-0f1fee6fda57)


---

## Notes

- The script automatically skips known system folders and lists for efficiency and accuracy.
- Handles SPO throttling gracefully with exponential backoff.
- No command-line parameters; all config is in the script.
- Review and update `$appID`, `$thumbprint`, `$tenant`, and `$inputFilePath` before use.

---

## Disclaimer

The script is provided AS IS, without warranty of any kind. Use at your own risk.

---

## Example

```powershell
.\Get-SPOItemPermissions.ps1
```

Make sure to update configuration variables in the script before running.

---

**Questions or contributions?** Please open an issue or pull request!

---

Let me know if you want this README tailored for the whole repo or just this script!
