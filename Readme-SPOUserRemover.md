# SPOUserRemover Documentation

## Overview
The SPOUserRemover.ps1 script is designed to facilitate the removal of users from SharePoint Online (SPO) sites efficiently. It supports multiple site collections and includes mechanisms to handle throttling, ensuring a smooth user management experience.

## Features
- **Multi-Site Support**: The script can handle user removal across multiple site collections seamlessly.
- **Intelligent Throttling**: Built-in mechanisms to manage throttling, ensuring compliance with SharePoint Online's limits.
- **Comprehensive Logging**: Logs all actions taken during script execution for review and auditing.
- **Three-Layer Access Removal**: Removes users at the site, list/library, and item levels to ensure complete access revocation.

## Prerequisites
- **PowerShell 5.1+**: Ensure your environment is running at least PowerShell version 5.1.
- **PnP.PowerShell**: Install the PnP.PowerShell module for SharePoint management.
- **Azure App Registration**: Register an app in Azure AD to authenticate with SharePoint Online.
- **X.509 Certificate**: Use an X.509 certificate for secure authentication.

## User Configuration
Example:
```powershell
# Configuring the script with necessary parameters
$Config = @{
    "SiteUrls" = @(
        "https://yourtenant.sharepoint.com/sites/Site1",
        "https://yourtenant.sharepoint.com/sites/Site2"
    );
    "UserEmail" = "user@example.com";
    "LogFilePath" = "C:\Logs\SPOUserRemover.log";
}
```

## How It Works
### Single Site Workflow
1. Connect to the given site.
2. Retrieve user information.
3. Remove users and log actions.

### Multi-Site Workflow
The script iterates through an array of site URLs and applies the same user removal process for each site.

## Usage Instructions
```powershell
# Running the SPOUserRemover Script
.\SPOUserRemover.ps1 -Configuration $Config
```

## Configuration Parameters
| Parameter       | Description                                      | Example                          |
|------------------|--------------------------------------------------|----------------------------------|
| SiteUrls         | List of site URLs to process                    | `https://tenant.sharepoint.com/sites/Site1`  |
| UserEmail        | The email of the user to remove                 | `user@example.com`               |
| LogFilePath      | Path to the log file produced                   | `C:\Logs\log.txt`              |

## Multi-Site Mode Details
When configured for multi-site, the script processes each site collectively, applying removal actions for the specified users.

## Throttling & Retry Mechanism
The script incorporates a back-off strategy that retries failed operations after a specified interval, to comply with SharePoint throttling policies.

## Logging Details
All actions taken by the script are logged, including timestamps, actions performed, and any errors encountered, stored in the specified log file.

## Troubleshooting Guide
1. **Issue**: User not found error.
   - **Solution**: Verify the email and ensure the user exists in the specified site.
2. **Issue**: Throttling errors.
   - **Solution**: Allow the script to pause and retry based on the throttling mechanism.
3. **Issue**: Insufficient permissions.
   - **Solution**: Ensure the account running the script has adequate permissions on the sites.
4. **Issue**: Logging not generating.
   - **Solution**: Check the log file path and ensure it is writable.
5. **Issue**: Incorrect site URL.
   - **Solution**: Double-check the site URLs for typos.
6. **Issue**: PowerShell module not installed.
   - **Solution**: Install required PowerShell modules like PnP.PowerShell.

## Performance Considerations
Monitor performance impacts when running the script against a large number of sites/users. Throttling may affect execution time.

## Notes
- Ensure to test the script in a non-production environment before applying it broadly.
- Review Azure AD permissions and ensure they are appropriate for the tasks being performed.

## Disclaimer
This script is provided "as is" without warranty of any kind. Use at your own risk and ensure compliance with organizational policies.