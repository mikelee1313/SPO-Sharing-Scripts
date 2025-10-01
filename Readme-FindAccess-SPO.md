# SharePoint Online Site Access Auditor

A comprehensive PowerShell script that audits SharePoint Online sites to determine user access across your entire Microsoft 365 tenant. This tool helps administrators understand who has access to what SharePoint resources through multiple access vectors.

## 🚀 Features

- **Comprehensive Access Detection**: Checks multiple access paths including:
  - Direct user permissions
  - Microsoft 365 Group membership  
  - Site ownership and site collection administration
  - SharePoint group membership
  - Entra ID (Azure AD) security group membership
  - Role assignments and permissions
  - "Everyone except external users" permissions

- **Enterprise-Grade Throttling Protection**: 
  - Exponential backoff retry logic with jitter
  - Respects SharePoint API rate limits and Retry-After headers
  - Configurable delays and retry attempts
  - Follows Microsoft's throttling guidance

- **Debug and Monitoring Capabilities**:
  - Detailed debug output (optional)
  - Comprehensive logging to file
  - Progress tracking with site-by-site status
  - Error handling and reporting

- **Customer Environment Ready**:
  - No hardcoded values outside configuration section
  - Easy tenant-specific customization
  - Dynamic tenant detection for EEEU patterns
  - Portable across different Microsoft 365 environments

## 📋 Prerequisites

### Required Modules
```powershell
Install-Module -Name PnP.PowerShell -Force -AllowClobber
```

### Required Permissions
The Entra ID app registration used for authentication must have the following **Application Permissions**:

- **SharePoint**:
  - `Sites.FullControl.All`
  - `User.Read.All`
  
- **Microsoft Graph**:
  - `Group.Read.All`
  - `User.Read.All`
  - `Directory.Read.All`

### Certificate-Based Authentication Setup
This script uses App-Only authentication with certificates. You'll need:
1. An Entra ID app registration
2. A certificate uploaded to the app registration
3. The certificate installed in the local certificate store
4. The certificate thumbprint

## ⚙️ Configuration

Before running the script, update these variables in the configuration section:

```powershell
#Configurable Settings
$t = 'YourTenantName' # Your tenant name (without .onmicrosoft.com)
$admin = 'admin@yourtenant.onmicrosoft.com'  # Your admin account
$tenant = '00000000-0000-0000-0000-000000000000'  # Your Tenant ID (GUID)
$appID = '00000000-0000-0000-0000-000000000000'   # Your App Registration ID
$thumbprint = 'ABCDEF1234567890ABCDEF1234567890ABCDEF12'  # Certificate thumbprint

# File Paths
$userListFile = 'C:\temp\users.txt'  # Path to file containing list of users

# Optional Feature Settings
$checkEEEU = $true  # Check for "Everyone except external users" permissions
$debug = $false     # Enable detailed debug output
```

## 📝 Input File Format

Create a text file with user principal names (one per line):

```
user1@yourdomain.com
user2@yourdomain.com
external.user@externaldomain.com
```

## 🏃‍♂️ Usage

### Basic Usage
```powershell
.\FindAccess-SPO.ps1
```

### With Debug Output
```powershell
# Method 1: Modify the $debug variable in the script
$debug = $true
.\FindAccess-SPO.ps1

# Method 2: Set before running
$debug = $true; .\FindAccess-SPO.ps1
```

### Custom Configuration
```powershell
# Disable EEEU checking for faster execution
$checkEEEU = $false
.\FindAccess-SPO.ps1
```

## 📊 Output

The script generates two output files with timestamps:

### 1. CSV Results File: `SiteUsers_[timestamp]_output.csv`
Contains audit results with columns:
- **SiteName**: Display name of the SharePoint site
- **URL**: Full URL of the SharePoint site  
- **User**: User principal name being checked
- **Owner**: Site owner
- **AccessType**: Type of access found

Example access types:
- `Direct Access - Member`
- `M365 Group Member`
- `Site Owner`
- `Site Collection Admin`
- `SharePoint Group: [GroupName]`
- `Entra ID Group: [GroupName]`
- `Everyone except external users`

### 2. Log File: `SiteUsers_[timestamp]_logfile.log`
Contains detailed execution log with:
- Timestamp for each operation
- Sites processed
- Users checked
- Errors encountered
- Performance metrics

## 🔧 Advanced Configuration

### Throttling Protection Settings
```powershell
$enableThrottlingProtection = $true    # Enable/disable throttling protection
$baseDelayBetweenSites = 2             # Seconds between sites
$baseDelayBetweenUsers = 1             # Seconds between users
$maxRetryAttempts = 5                  # Max retries for failed operations
$baseRetryDelay = 30                   # Base retry delay (exponential backoff)
```

### Performance Tuning
- **Large tenants**: Increase `$baseDelayBetweenSites` to reduce API load
- **Small tenants**: Decrease delays for faster execution
- **Rate limiting issues**: Increase `$baseRetryDelay` and `$maxRetryAttempts`

## 🚨 Troubleshooting

### Common Issues

**1. Authentication Failures**
```
Error: AADSTS70002: Error validating credentials
```
- Verify certificate is installed correctly
- Check thumbprint matches exactly
- Ensure app registration has required permissions
- Verify tenant ID is correct

**2. Permission Errors**
```
Access denied. You do not have permission to perform this action
```
- Check app registration has required SharePoint permissions
- Verify admin consent has been granted
- Ensure certificate-based authentication is properly configured

**3. Throttling Issues**
```
Request rate is too high
```
- Script automatically handles throttling with retry logic
- Increase `$baseDelayBetweenSites` if issues persist
- Enable debug mode to see throttling protection in action

**4. EEEU Detection Issues**
```
EEEU not present in site users - skipping EEEU check
```
- Verify `$t` (tenant name) is configured correctly
- Check that sites actually have EEEU permissions configured
- Set `$debug = $true` to see EEEU pattern being used

### Debug Mode
Enable debug output for detailed troubleshooting:
```powershell
$debug = $true
```

This will show:
- Connection attempts and status
- Site processing progress  
- User access checking details
- EEEU pattern detection
- Throttling protection activation
- API call details and timing

## 📋 Best Practices

1. **Start Small**: Test with a small user list first
2. **Monitor Performance**: Use debug mode to understand timing
3. **Schedule Wisely**: Run during off-peak hours for large tenants
4. **Regular Audits**: Schedule periodic access reviews
5. **Backup Results**: Archive CSV outputs for compliance

## 🔒 Security Considerations

- Store certificates securely
- Use dedicated service accounts
- Limit app registration permissions to minimum required
- Regularly review and rotate certificates
- Monitor script execution logs

## 📜 License

This script is provided as-is under the MIT License. See the disclaimer in the script header for full terms.

## 🤝 Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Test thoroughly with your changes
4. Submit a pull request with detailed description

## 📞 Support

For issues and questions:
1. Check the troubleshooting section above
2. Enable debug mode for detailed error information
3. Review the execution log file
4. Submit an issue with logs and configuration details (redact sensitive information)

---

**Note**: This script performs read-only operations on SharePoint Online. It does not modify any site permissions or user access.
