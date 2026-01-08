<#
.SYNOPSIS
    SharePoint Online Site Access Auditor - Checks user access across all SharePoint sites in a tenant.

.DESCRIPTION
    This script audits SharePoint Online sites to determine if specified users have access to each site.
    It uses PnP PowerShell with App-Only Authentication and checks multiple access vectors including:
    - Direct user permissions
    - Microsoft 365 Group membership
    - Site ownership
    - Site collection administration
    - SharePoint group membership
    - Entra ID (Azure AD) security group membership
    - Role assignments and permissions
    - "Everyone except external users" permissions (optional)
    
    The script includes comprehensive throttling protection with exponential backoff retry logic
    to handle SharePoint API rate limits gracefully.

.PARAMETER appID
    The Application (Client) ID of the Entra ID app registration used for authentication.

.PARAMETER thumbprint
    The certificate thumbprint for the certificate used in App-Only authentication.

.PARAMETER tenant
    The Tenant ID (GUID) of your Microsoft 365 tenant.

.PARAMETER t
    The tenant name (without .onmicrosoft.com) for building SharePoint URLs.

.PARAMETER admin
    The admin account UPN for reference (not used for authentication in App-Only mode).

.PARAMETER users
    Array of user principal names (UPNs) to check for access. Loaded from a text file.

.PARAMETER checkEEEU
    Boolean flag to enable/disable checking for "Everyone except external users" permissions.
    Default: $true

.PARAMETER debug
    Boolean flag to enable/disable detailed debug output during script execution.
    Default: $false

.PARAMETER enableThrottlingProtection
    Boolean flag to enable/disable throttling protection and retry logic.
    Default: $true (recommended)

.PARAMETER baseDelayBetweenSites
    Base delay in seconds between processing different SharePoint sites.
    Default: 2 seconds

.PARAMETER baseDelayBetweenUsers
    Base delay in seconds between processing different users within the same site.
    Default: 1 second

.PARAMETER maxRetryAttempts
    Maximum number of retry attempts when SharePoint API throttling is encountered.
    Default: 5

.PARAMETER baseRetryDelay
    Base retry delay in seconds when throttling occurs. Uses exponential backoff.
    Default: 30 seconds

.INPUTS
    Text file containing user principal names (one per line) specified in $users variable.

.OUTPUTS
    CSV file: SiteUsers_[timestamp]_output.csv - Contains audit results with columns:
    - SiteName: Display name of the SharePoint site
    - URL: Full URL of the SharePoint site
    - User: User principal name being checked
    - Owner: Site owner
    - AccessType: Type of access found (Direct Access, M365 Group Member, Site Owner, etc.)
    
    Log file: SiteUsers_[timestamp]_logfile.log - Contains detailed execution log

.EXAMPLE
    # Basic usage with default settings
    .\FindAccess-SPO.ps1
    
    # Run with debug output enabled
    $debug = $true
    .\FindAccess-SPO.ps1
    
    # Run without throttling protection (not recommended)
    $enableThrottlingProtection = $false
    .\FindAccess-SPO.ps1

.NOTES
    File Name      : FindAccess-SPO.ps1
    Author         : Mike Lee
    Prerequisite   : PnP PowerShell module, App-Only authentication setup
    Created Date   : 10/1/2025
    Updated:       : 12/2/2025
    Version        : 2.0 (Converted to PnP PowerShell with App-Only Authentication)

Disclaimer: The sample scripts are provided AS IS without warranty  
        of any kind. Microsoft further disclaims all implied warranties including,  
        without limitation, any implied warranties of merchantability or of fitness for 
        a particular purpose. The entire risk arising out of the use or performance of  
        the sample scripts and documentation remains with you. In no event shall 
        Microsoft, its authors, or anyone else involved in the creation, production, or 
        delivery of the scripts be liable for any damages whatsoever (including, 
        without limitation, damages for loss of business profits, business interruption, 
        loss of business information, or other pecuniary loss) arising out of the use 
        of or inability to use the sample scripts or documentation, even if Microsoft 
        has been advised of the possibility of such damages.

    REQUIREMENTS:
    - PnP PowerShell module installed
    - Entra ID App Registration with appropriate SharePoint permissions
    - Certificate-based authentication configured
    - SharePoint Administrator or Global Administrator permissions
    - Users.txt file with list of UPNs to check
    
    PERMISSIONS REQUIRED:
    The Entra ID app registration needs the following permissions:
    - SharePoint: Sites.FullControl.All (Application)
    - Graph: Sites.Read.All (Application)
    

.LINK
    https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets
    https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs

.COMPONENT
    PnP PowerShell, SharePoint Online, Microsoft 365, Entra ID

.FUNCTIONALITY
    SharePoint Online access auditing, user permission analysis, compliance reporting
#>

# =================================================================================================
# USER CONFIGURATION - Update the variables in this section
# =================================================================================================

# App-Only Authentication Settings
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536" # This is your Entra App ID
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082" # This is certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3" # This is your Tenant ID

#Configurable Settings
$t = 'M365CPI13246019' # < - Your Tenant Name Here
$admin = 'admin@M365CPI13246019.onmicrosoft.com'  # <- Your Admin Account Here

#A Simple List of users by UPN
$users = Get-Content 'C:\temp\users.txt'

# Optional Feature Settings
$checkEEEU = $true  # Set to $true to check for "Everyone except external users" permissions, $false to skip
$debug = $false  # Set to $true for detailed debug output, $false for minimal output

# Throttling Protection Settings
$enableThrottlingProtection = $true  # Set to $false to disable throttling protection (not recommended)
$baseDelayBetweenSites = 2  # Base delay in seconds between processing sites
$baseDelayBetweenUsers = 1  # Base delay in seconds between processing users within a site
$maxRetryAttempts = 5  # Maximum number of retry attempts when throttled
$baseRetryDelay = 30  # Base retry delay in seconds (will be increased with exponential backoff)

# =================================================================================================
# END OF USER CONFIGURATION
# =================================================================================================


# =================================================================================================
# FUNCTION DEFINITIONS
# =================================================================================================

#This is the logging function - MUST be defined first as other functions depend on it
Function Write-LogEntry {
    param(
        [string] $LogName ,
        [string] $LogEntryText
    )
    if ($LogName -NotLike $Null) {
        # log the date and time in the text file along with the data passed
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : $LogEntryText" | Out-File -FilePath $LogName -append;
    }
}

# Debug-aware output functions
Function Write-DebugInfo {
    param(
        [string] $Message,
        [string] $ForegroundColor = "Gray",
        [switch] $AlwaysLog
    )
    
    if ($debug) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
    
    # Always log debug messages when debug is enabled, or when AlwaysLog is specified
    if ($debug -or $AlwaysLog) {
        Write-LogEntry -LogName:$log -LogEntryText $Message
    }
}

Function Write-InfoMessage {
    param(
        [string] $Message,
        [string] $ForegroundColor = "White",
        [switch] $AlwaysShow
    )
    
    # Show info messages when debug is enabled or when AlwaysShow is specified
    if ($debug -or $AlwaysShow) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
    
    # Always log info messages
    Write-LogEntry -LogName:$log -LogEntryText $Message
}

Function Write-StatusMessage {
    param(
        [string] $Message,
        [string] $ForegroundColor = "Yellow",
        [switch] $Force
    )
    
    # Always show critical status messages, or when forced
    Write-Host $Message -ForegroundColor $ForegroundColor
    
    # Always log status messages
    Write-LogEntry -LogName:$log -LogEntryText $Message
}

# Function to handle SharePoint throttling with exponential backoff and retry logic
Function Invoke-PnPCommandWithThrottling {
    param(
        [scriptblock] $Command,
        [string] $OperationDescription = "SharePoint operation",
        [int] $MaxRetries = $maxRetryAttempts,
        [int] $BaseDelay = $baseRetryDelay
    )
    
    if (-not $enableThrottlingProtection) {
        # If throttling protection is disabled, execute command directly
        return & $Command
    }
    
    $attempt = 0
    $delay = $BaseDelay
    
    while ($attempt -lt $MaxRetries) {
        try {
            $attempt++
            Write-DebugInfo "    Executing: $OperationDescription (attempt $attempt/$MaxRetries)"
            
            $result = & $Command
            
            # If we get here, the command succeeded
            if ($attempt -gt 1) {
                Write-StatusMessage "    ✓ $OperationDescription succeeded after $attempt attempts" -ForegroundColor Green -Force
            }
            return $result
        }
        catch {
            $statusCode = $null
            $retryAfter = $null
            
            # Try to extract HTTP status code and Retry-After header from the exception
            if ($_.Exception.Message -match "429" -or $_.Exception.Message -match "Too Many Requests") {
                $statusCode = 429
                Write-StatusMessage "    ⚠️ HTTP 429 (Too Many Requests) detected for: $OperationDescription" -ForegroundColor Yellow -Force
            }
            elseif ($_.Exception.Message -match "503" -or $_.Exception.Message -match "Server Too Busy") {
                $statusCode = 503
                Write-StatusMessage "    ⚠️ HTTP 503 (Server Too Busy) detected for: $OperationDescription" -ForegroundColor Yellow -Force
            }
            elseif ($_.Exception.Message -match "throttl" -or $_.Exception.Message -match "rate limit") {
                Write-StatusMessage "    ⚠️ Throttling detected for: $OperationDescription" -ForegroundColor Yellow -Force
                $statusCode = 429  # Treat as throttling
            }
            
            # If this is a throttling error and we have retries left
            if (($statusCode -eq 429 -or $statusCode -eq 503) -and $attempt -lt $MaxRetries) {
                # Try to extract Retry-After header value from exception message
                if ($_.Exception.Message -match "Retry-After[:\s]+(\d+)") {
                    $retryAfter = [int]$matches[1]
                    Write-DebugInfo "    📋 Retry-After header indicates waiting $retryAfter seconds" -ForegroundColor Cyan
                }
                
                # Calculate delay: use Retry-After if available, otherwise exponential backoff
                if ($retryAfter -and $retryAfter -gt 0) {
                    $waitTime = $retryAfter
                }
                else {
                    # Exponential backoff with jitter
                    $exponentialDelay = $delay * [Math]::Pow(2, $attempt - 1)
                    $jitter = Get-Random -Minimum 0 -Maximum 10
                    $waitTime = $exponentialDelay + $jitter
                }
                
                Write-StatusMessage "    ⏳ Throttling detected. Waiting $waitTime seconds before retry $($attempt + 1)/$MaxRetries..." -ForegroundColor Yellow -Force
                Write-LogEntry -LogName:$log -LogEntryText "Throttling detected for $OperationDescription. Waiting $waitTime seconds before retry $($attempt + 1)/$MaxRetries"
                
                Start-Sleep -Seconds $waitTime
                continue
            }
            else {
                # Either not a throttling error, or we've exhausted retries
                if ($attempt -ge $MaxRetries) {
                    Write-StatusMessage "    ❌ Maximum retry attempts ($MaxRetries) exceeded for: $OperationDescription" -ForegroundColor Red -Force
                    Write-LogEntry -LogName:$log -LogEntryText "Maximum retry attempts exceeded for $OperationDescription. Last error: $($_.Exception.Message)"
                }
                else {
                    Write-DebugInfo "    ❌ Non-throttling error for: $OperationDescription - $($_.Exception.Message)" -ForegroundColor Red -AlwaysLog
                }
                throw
            }
        }
    }
}

# Function to add intelligent delays between operations to prevent overwhelming SharePoint
Function Add-ThrottlingDelay {
    param(
        [string] $DelayType = "user",  # "site" or "user"
        [string] $Description = ""
    )
    
    if (-not $enableThrottlingProtection) {
        return
    }
    
    $delaySeconds = 0
    
    switch ($DelayType.ToLower()) {
        "site" {
            $delaySeconds = $baseDelayBetweenSites
        }
        "user" {
            $delaySeconds = $baseDelayBetweenUsers
        }
        default {
            $delaySeconds = 1
        }
    }
    
    if ($delaySeconds -gt 0) {
        Write-DebugInfo "    ⏱️ Adding $delaySeconds second delay$($Description)"
        Start-Sleep -Seconds $delaySeconds
    }
}

# Function to check if user is member of Entra ID group using PnP commands
Function Test-EntraGroupMembership {
    param(
        [string] $UserPrincipalName,
        [string] $GroupId,
        [string] $GroupDisplayName
    )
    
    try {
        # Try to get group members using PnP commands
        # First try as Microsoft 365 Group
        $groupMembers = Get-PnPMicrosoft365GroupMember -Identity $GroupId -ErrorAction SilentlyContinue
        if ($groupMembers) {
            $isMember = $groupMembers | Where-Object {
                $_.UserPrincipalName -eq $UserPrincipalName -or 
                $_.Mail -eq $UserPrincipalName
            }
            if ($isMember) {
                return $true
            }
        }
        
        # Also check for group owners (Microsoft 365 Groups)
        $groupOwners = Get-PnPMicrosoft365GroupOwner -Identity $GroupId -ErrorAction SilentlyContinue
        if ($groupOwners) {
            $isOwner = $groupOwners | Where-Object {
                $_.UserPrincipalName -eq $UserPrincipalName -or 
                $_.Mail -eq $UserPrincipalName
            }
            if ($isOwner) {
                Write-DebugInfo "User is an OWNER of Microsoft 365 group '$GroupDisplayName'" -ForegroundColor Magenta
                return $true
            }
        }
        
        # If that fails, try as Azure AD Group using Get-PnPAzureADGroup
        $azureGroup = Get-PnPAzureADGroup -Identity $GroupId -ErrorAction SilentlyContinue
        if ($azureGroup) {
            $azureGroupMembers = Get-PnPAzureADGroupMember -Identity $GroupId -ErrorAction SilentlyContinue
            if ($azureGroupMembers) {
                $isMember = $azureGroupMembers | Where-Object {
                    $_.UserPrincipalName -eq $UserPrincipalName -or 
                    $_.Mail -eq $UserPrincipalName
                }
                return ($null -ne $isMember)
            }
        }
        
        return $false
    }
    catch {
        Write-DebugInfo "Could not check membership for group '$GroupDisplayName': $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $false
    }
}

# =================================================================================================
# CONNECTION AND INITIALIZATION
# =================================================================================================

#Connect to SharePoint Online using PnP PowerShell with App-Only Authentication
try {
    Write-StatusMessage "Connecting to SharePoint Online using App-Only Authentication..." -ForegroundColor Green
    Connect-PnPOnline -Url "https://$t-admin.sharepoint.com" -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
    Write-StatusMessage "Successfully connected to SharePoint Online" -ForegroundColor Green
    
    if ($enableThrottlingProtection) {
        Write-DebugInfo "Throttling protection is ENABLED with the following settings:" -ForegroundColor Green
        Write-DebugInfo "  - Base delay between sites: $baseDelayBetweenSites seconds" -ForegroundColor Gray
        Write-DebugInfo "  - Base delay between users: $baseDelayBetweenUsers seconds" -ForegroundColor Gray
        Write-DebugInfo "  - Max retry attempts: $maxRetryAttempts" -ForegroundColor Gray
        Write-DebugInfo "  - Base retry delay: $baseRetryDelay seconds" -ForegroundColor Gray
    }
    else {
        Write-DebugInfo "⚠️ WARNING: Throttling protection is DISABLED" -ForegroundColor Yellow
    }
    
    # Show debug mode status
    if ($debug) {
        Write-DebugInfo "DEBUG MODE: ENABLED - Detailed output will be shown" -ForegroundColor Cyan
    }
    else {
        Write-DebugInfo "DEBUG MODE: DISABLED - Minimal output will be shown" -ForegroundColor Gray
    }
}
catch {
    Write-StatusMessage "Failed to connect to SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Note: PnP PowerShell can handle Entra ID groups directly, no separate Graph connection needed

#Initialize Parameters - Do not change
$user = @()
$date = @()
$outputfile = @()
$log = @()
$date = Get-Date -Format yyyy-MM-dd_HH-mm-ss
$firstWrite = $true

#OutPut and Log Files
$outputfile = "$env:TEMP\" + 'SiteUsers_' + $date + "output.csv"
$log = "$env:TEMP\" + 'SiteUsers_' + $date + '_' + "logfile.log"

#Get All Sites that are not Group Connected and exclude system/service sites
$sites = Get-PnPTenantSite -includeOneDriveSites | Where-Object {
    $_.Template -ne 'RedirectSite#0' -and
    $_.Template -notlike 'SRCHCEN*' -and
    $_.Template -notlike 'SRCHCENTERLITE*' -and
    $_.Template -notlike 'SPSMSITEHOST*' -and
    $_.Template -notlike 'APPCATALOG*' -and
    $_.Template -notlike 'REDIRECTSITE*'
}

# =================================================================================================
# MAIN PROCESSING LOOP
# =================================================================================================

#Starting Loop for All Sites
foreach ($site in $sites) {
    
    # Add delay between sites to prevent overwhelming SharePoint
    Add-ThrottlingDelay -DelayType "site" -Description " between sites to prevent throttling"
    
    # Initialize site-specific output array
    $siteOutput = @()
    
    Write-StatusMessage "Processing site: $($site.Title) ($($site.url))" -ForegroundColor Yellow
    Write-LogEntry -LogName:$Log -LogEntryText "Starting processing for site: $($site.Title) ($($site.url))"

    #Starting Loop of all users in the list provided
    foreach ($user in $users) {
        
        # Add delay between users to prevent overwhelming SharePoint
        Add-ThrottlingDelay -DelayType "user" -Description " between users to prevent throttling"
        
        #Connect to all sites as the Users provided
        Write-DebugInfo "Attempting to GET '$user' on SITE '$($site.url)'" -ForegroundColor Green
        Write-LogEntry -LogName:$Log -LogEntryText "Attempting to GET '$user' on SITE '$($site.url)'"
        
        try {
            # Connect to the specific site to check user permissions with throttling protection
            Invoke-PnPCommandWithThrottling -Command {
                Connect-PnPOnline -Url $site.Url -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
            } -OperationDescription "Connect to site $($site.Url)" | Out-Null
            
            $userFound = $false
            $accessType = ""
            $groupMemberships = @()
            
            # Check 1: Direct user access
            $SiteMember = Invoke-PnPCommandWithThrottling -Command {
                Get-PnPUser -Identity $user -ErrorAction SilentlyContinue
            } -OperationDescription "Get user $user from site"
            
            if ($SiteMember) {
                $userFound = $true
                $accessType = "Direct Access"
                Write-DebugInfo "Found $user with direct access on '$($site.url)'" -ForegroundColor Cyan
            }
            
            # Check 1.5: Microsoft 365 Group-connected site membership
            if ($site.GroupId -and $site.GroupId -ne "00000000-0000-0000-0000-000000000000") {
                Write-DebugInfo "Site is connected to Microsoft 365 Group (ID: $($site.GroupId)). Checking group membership..." -ForegroundColor Yellow
                
                try {
                    # Check if user is a member of the connected Microsoft 365 group
                    $groupMembers = Invoke-PnPCommandWithThrottling -Command {
                        Get-PnPMicrosoft365GroupMember -Identity $site.GroupId -ErrorAction SilentlyContinue
                    } -OperationDescription "Get M365 group members for $($site.GroupId)"
                    if ($groupMembers) {
                        $isMember = $groupMembers | Where-Object {
                            $_.UserPrincipalName -eq $user -or 
                            $_.Mail -eq $user
                        }
                        if ($isMember) {
                            $userFound = $true
                            $accessType += "; M365 Group Member: $($site.Title)"
                            $groupMemberships += "M365 Group: $($site.Title)"
                            Write-DebugInfo "✓ Found $user as member of connected Microsoft 365 group for '$($site.url)'" -ForegroundColor Green
                        }
                    }
                    
                    # Also check if user is an owner of the connected Microsoft 365 group
                    $groupOwners = Invoke-PnPCommandWithThrottling -Command {
                        Get-PnPMicrosoft365GroupOwner -Identity $site.GroupId -ErrorAction SilentlyContinue
                    } -OperationDescription "Get M365 group owners for $($site.GroupId)"
                    if ($groupOwners) {
                        $isOwner = $groupOwners | Where-Object {
                            $_.UserPrincipalName -eq $user -or 
                            $_.Mail -eq $user
                        }
                        if ($isOwner) {
                            $userFound = $true
                            $accessType += "; M365 Group Owner: $($site.Title)"
                            $groupMemberships += "M365 Group Owner: $($site.Title)"
                            Write-DebugInfo "✓ Found $user as OWNER of connected Microsoft 365 group for '$($site.url)'" -ForegroundColor Magenta
                        }
                    }
                }
                catch {
                    Write-DebugInfo "Could not check Microsoft 365 group membership: $($_.Exception.Message)" -ForegroundColor DarkYellow
                }
            }
            else {
                Write-DebugInfo "Site is not connected to a Microsoft 365 Group" -ForegroundColor DarkGray
            }
            
            # Check 2: Site Owner
            if ($site.Owner -eq $user) {
                $userFound = $true
                $accessType += "; Site Owner"
                Write-DebugInfo "Found $user as site owner on '$($site.url)'" -ForegroundColor Cyan
            }
            
            # Check 3: Site Collection Administrator
            try {
                Write-DebugInfo "Checking Site Collection Administrators..." -ForegroundColor DarkYellow
                $siteCollectionAdmins = Invoke-PnPCommandWithThrottling -Command {
                    Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
                } -OperationDescription "Get site collection admins"
                
                if ($siteCollectionAdmins) {
                    Write-DebugInfo "Found $($siteCollectionAdmins.Count) site collection admin(s)" -ForegroundColor DarkYellow
                    
                    $isAdmin = $false
                    foreach ($admin in $siteCollectionAdmins) {
                        # Try to resolve the claim to get the actual user principal name
                        $resolvedUser = $null
                        
                        # For tenant claims (c:0t.c|tenant|guid), try to get the user details
                        if ($admin.LoginName -like "c:0t.c|tenant|*") {
                            try {
                                $admin.LoginName -replace "c:0t.c\|tenant\|", ""
                                # Try to get user by ID - this might help resolve the actual UPN
                                $resolvedUser = Get-PnPUser -Identity $admin.LoginName -ErrorAction SilentlyContinue
                                if ($resolvedUser) {
                                    $actualUserPrincipal = $resolvedUser.UserPrincipalName
                                    $actualEmail = $resolvedUser.Email
                                    
                                    Write-DebugInfo "  - Resolved: $($admin.LoginName) -> $actualUserPrincipal ($actualEmail)" -ForegroundColor DarkCyan
                                    
                                    # Check if this matches our target user
                                    if ($actualUserPrincipal -eq $user -or $actualEmail -eq $user) {
                                        $isAdmin = $true
                                        break
                                    }
                                }
                            }
                            catch {
                                Write-DebugInfo "  - Could not resolve: $($admin.LoginName)" -ForegroundColor DarkGray
                            }
                        }
                        # For federated claims (c:0o.c|federateddirectoryclaimprovider|guid_o), different handling
                        elseif ($admin.LoginName -like "c:0o.c|federateddirectoryclaimprovider|*") {
                            try {
                                $resolvedUser = Get-PnPUser -Identity $admin.LoginName -ErrorAction SilentlyContinue
                                if ($resolvedUser) {
                                    $actualUserPrincipal = $resolvedUser.UserPrincipalName
                                    $actualEmail = $resolvedUser.Email
                                    
                                    Write-DebugInfo "  - Resolved: $($admin.LoginName) -> $actualUserPrincipal ($actualEmail)" -ForegroundColor DarkCyan
                                    
                                    # Check if this matches our target user
                                    if ($actualUserPrincipal -eq $user -or $actualEmail -eq $user) {
                                        $isAdmin = $true
                                        break
                                    }
                                }
                            }
                            catch {
                                Write-DebugInfo "  - Could not resolve: $($admin.LoginName)" -ForegroundColor DarkGray
                            }
                        }
                        # For regular user formats
                        else {
                            Write-DebugInfo "  - Regular format: $($admin.LoginName) ($($admin.UserPrincipalName)) ($($admin.Email))" -ForegroundColor DarkCyan
                            
                            if ($admin.LoginName -eq $user -or 
                                $admin.Email -eq $user -or 
                                $admin.UserPrincipalName -eq $user -or
                                $admin.LoginName -like "*$user*" -or
                                $admin.UserPrincipalName -like "*$user*") {
                                $isAdmin = $true
                                break
                            }
                        }
                    }
                    
                    if ($isAdmin) {
                        $userFound = $true
                        $accessType += "; Site Collection Admin"
                        Write-DebugInfo "✓ Found $user as Site Collection Administrator on '$($site.url)'" -ForegroundColor Cyan
                    }
                    else {
                        Write-DebugInfo "User not found in site collection admins list after claim resolution" -ForegroundColor DarkYellow
                    }
                }
                else {
                    Write-DebugInfo "No site collection administrators returned from Get-PnPSiteCollectionAdmin" -ForegroundColor DarkYellow
                }
            }
            catch {
                Write-DebugInfo "Error checking site collection admins: $($_.Exception.Message)" -ForegroundColor Red
                # If we can't get site collection admins, try alternative method
                try {
                    Write-DebugInfo "Trying alternative method - checking Full Control permissions..." -ForegroundColor DarkYellow
                    # Alternative: Check if user has Full Control at site level
                    $sitePermissions = Invoke-PnPCommandWithThrottling -Command {
                        Get-PnPRoleAssignment -ErrorAction SilentlyContinue
                    } -OperationDescription "Get site role assignments for Full Control check"
                    foreach ($permission in $sitePermissions) {
                        if (($permission.Member.LoginName -eq $user -or 
                                $permission.Member.Email -eq $user -or
                                $permission.Member.LoginName -like "*$user*") -and 
                            $permission.RoleDefinitionBindings.Name -contains "Full Control") {
                            $userFound = $true
                            $accessType += "; Full Control (Site Admin)"
                            Write-DebugInfo "✓ Found $user with Full Control permissions on '$($site.url)'" -ForegroundColor Cyan
                            break
                        }
                    }
                }
                catch {
                    Write-DebugInfo "Alternative method also failed: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            
            # Check 4: SharePoint Groups membership
            $siteGroups = Invoke-PnPCommandWithThrottling -Command {
                Get-PnPGroup -ErrorAction SilentlyContinue
            } -OperationDescription "Get SharePoint groups"
            
            foreach ($group in $siteGroups) {
                try {
                    # Skip groups with "Limited Access" as they don't represent meaningful permissions
                    if ($group.Title -like "*Limited Access*") {
                        continue
                    }
                    
                    $groupMembers = Invoke-PnPCommandWithThrottling -Command {
                        Get-PnPGroupMember -Identity $group.Title -ErrorAction SilentlyContinue
                    } -OperationDescription "Get members of SharePoint group $($group.Title)"
                    if ($groupMembers | Where-Object { $_.LoginName -eq $user -or $_.Email -eq $user }) {
                        $userFound = $true
                        $accessType += "; SharePoint Group: $($group.Title)"
                        $groupMemberships += $group.Title
                        Write-DebugInfo "Found $user in SharePoint group '$($group.Title)' on '$($site.url)'" -ForegroundColor Cyan
                    }
                }
                catch {
                    # Skip groups we can't access
                }
            }
            
            # Check 5: Check role assignments (permissions) including Entra ID groups
            try {
                # Use Get-PnPUser instead of Get-PnPRoleAssignment for better compatibility
                Write-DebugInfo "Checking site users for Entra ID groups and direct assignments..." -ForegroundColor Yellow
                $siteUsers = Invoke-PnPCommandWithThrottling -Command {
                    Get-PnPUser -ErrorAction SilentlyContinue
                } -OperationDescription "Get site users for Entra ID group check"
                
                if ($siteUsers) {
                    foreach ($siteUser in $siteUsers) {
                        # Check for direct user assignment
                        if ($siteUser.LoginName -eq $user -or $siteUser.Email -eq $user -or $siteUser.UserPrincipalName -eq $user) {
                            $userFound = $true
                            $accessType += "; Direct User Assignment"
                            Write-DebugInfo "Found $user with direct user assignment on '$($site.url)'" -ForegroundColor Cyan
                        }
                        
                        # Check if this is an Entra ID group (PrincipalType = SecurityGroup)
                        if ($siteUser.PrincipalType -eq "SecurityGroup") {
                            try {
                                # Get Entra ID group information
                                $groupLoginName = $siteUser.LoginName
                                $groupTitle = $siteUser.Title
                                
                                # Extract group ID from LoginName if it's an Entra ID group
                                if ($groupLoginName -like "c:0t.c|tenant|*") {
                                    $groupId = $groupLoginName -replace "c:0t.c\|tenant\|", ""
                                    
                                    Write-DebugInfo "Checking Entra ID group membership for '$groupTitle'..." -ForegroundColor Yellow
                                    
                                    # Check if user is member of this Entra ID group
                                    $isMember = Test-EntraGroupMembership -UserPrincipalName $user -GroupId $groupId -GroupDisplayName $groupTitle
                                    
                                    if ($isMember) {
                                        $userFound = $true
                                        $accessType += "; Entra ID Group: $groupTitle"
                                        $groupMemberships += "Entra ID: $groupTitle"
                                        Write-DebugInfo "✓ Found $user in Entra ID group '$groupTitle' on '$($site.url)'" -ForegroundColor Green
                                    }
                                    else {
                                        Write-DebugInfo "✗ User $user is NOT in Entra ID group '$groupTitle' on '$($site.url)'" -ForegroundColor DarkGray
                                    }
                                }
                                else {
                                    # This might be a SharePoint group, mark for potential access
                                    Write-DebugInfo "Found security group '$groupTitle' (non-Entra ID pattern: $groupLoginName)" -ForegroundColor Yellow
                                }
                            }
                            catch {
                                Write-DebugInfo "Could not check Entra ID group membership for '$groupTitle'" -ForegroundColor DarkYellow
                            }
                        }
                    }
                }
                else {
                    Write-DebugInfo "No site users found" -ForegroundColor DarkGray
                }
            }
            catch {
                Write-DebugInfo "Error getting site users: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            # Check 6: Check for "Everyone except external users" permissions (optional)
            if ($checkEEEU) {
                try {
                    Write-DebugInfo "Checking for 'Everyone except external users' permissions..." -ForegroundColor Yellow
                    $everyoneExceptExternalFound = $false
                    $EEEU = '*spo-grid-all-users*'
                
                    # First, check if EEEU exists in the site users list
                    $eeeuInSiteUsers = $false
                    try {
                        $siteUsers = Invoke-PnPCommandWithThrottling -Command {
                            Get-PnPUser -WithRightsAssigned -ErrorAction SilentlyContinue
                        } -OperationDescription "Get site users with rights for EEEU check"
                        $eeeuUser = $siteUsers | Where-Object { $_.LoginName -like $EEEU }
                        if ($eeeuUser) {
                            $eeeuInSiteUsers = $true
                            Write-DebugInfo "    ✓ Found EEEU in site users list: $($eeeuUser.LoginName)" -ForegroundColor DarkCyan
                        }
                        else {
                            Write-DebugInfo "    EEEU not found in site users list" -ForegroundColor DarkGray
                        }
                    }
                    catch {
                        Write-Host "    Could not check site users for EEEU: $($_.Exception.Message)" -ForegroundColor DarkYellow
                    }
                
                    # Only proceed if EEEU exists in site users
                    if ($eeeuInSiteUsers) {
                        try {
                            # Get role assignments
                            $Permissions = Invoke-PnPCommandWithThrottling -Command {
                                Get-PnPProperty -ClientObject (Get-PnPWeb) -Property RoleAssignments -ErrorAction SilentlyContinue
                            } -OperationDescription "Get web role assignments for EEEU check"
                        
                            if ($Permissions -and $Permissions.Count -gt 0) {
                                Write-DebugInfo "    Found $($Permissions.Count) role assignments to check for EEEU" -ForegroundColor DarkCyan
                                $directEEEUFound = $false
                        
                                # First, check for direct EEEU permissions
                                foreach ($RoleAssignment in $Permissions) {
                                    try {
                                        # Get role assignment properties
                                        Invoke-PnPCommandWithThrottling -Command {
                                            Get-PnPProperty -ClientObject $RoleAssignment -Property Member -ErrorAction SilentlyContinue | Out-Null
                                            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings -ErrorAction SilentlyContinue | Out-Null
                                        } -OperationDescription "Get role assignment properties for EEEU direct check" | Out-Null

                                        # Check for direct EEEU with meaningful permissions
                                        if ($RoleAssignment.Member.LoginName -like $EEEU -and $RoleAssignment.RoleDefinitionBindings.Name -ne 'Limited Access') {
                                            Write-DebugInfo "    ✓ Found EEEU with direct meaningful permissions" -ForegroundColor Green
                                            Write-DebugInfo "    ✓ EEEU roles: $($RoleAssignment.RoleDefinitionBindings.Name -join ', ')" -ForegroundColor Green
                                            $directEEEUFound = $true
                                        
                                            # Only grant access if user is internal
                                            if ($user -like "*@$t.onmicrosoft.com" -or $user -like "*@*.onmicrosoft.com") {
                                                $userFound = $true
                                                $accessType += "; Everyone except external users"
                                                $everyoneExceptExternalFound = $true
                                                Write-DebugInfo "✓ Found $user has access via 'Everyone except external users' (direct) on '$($site.url)'" -ForegroundColor Green
                                                break
                                            }
                                            else {
                                                Write-Host "    User is external, not granting access via 'Everyone except external users'" -ForegroundColor DarkYellow
                                            }
                                        }
                                    }
                                    catch {
                                        # Skip role assignments that fail to load
                                        Write-DebugInfo "    Warning: Could not process role assignment: $($_.Exception.Message)" -ForegroundColor DarkYellow
                                    }
                                }
                            
                                # If direct EEEU not found, check SharePoint groups for EEEU membership
                                if (-not $directEEEUFound -and -not $everyoneExceptExternalFound) {
                                    Write-DebugInfo "    No direct EEEU permissions found, checking SharePoint groups..." -ForegroundColor DarkCyan
                                
                                    foreach ($RoleAssignment in $Permissions) {
                                        try {
                                            # Get role assignment properties
                                            Invoke-PnPCommandWithThrottling -Command {
                                                Get-PnPProperty -ClientObject $RoleAssignment -Property Member -ErrorAction SilentlyContinue | Out-Null
                                                Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings -ErrorAction SilentlyContinue | Out-Null
                                            } -OperationDescription "Get role assignment properties for EEEU SharePoint group check" | Out-Null

                                            # Check if this is a SharePoint group with meaningful permissions
                                            if ($RoleAssignment.Member.PrincipalType -eq "SharePointGroup" -and 
                                                $RoleAssignment.RoleDefinitionBindings.Name -ne 'Limited Access') {
                            
                                                $groupTitle = $RoleAssignment.Member.Title
                            
                                                # Skip groups with "Limited Access" in the name as they don't represent meaningful permissions
                                                if ($groupTitle -like "*Limited Access*") {
                                                    continue
                                                }
                            
                                                Write-DebugInfo "    Checking SharePoint group '$groupTitle' for EEEU membership..." -ForegroundColor DarkCyan
                                            
                                                try {
                                                    # Check if EEEU is a member of this SharePoint group
                                                    $groupMembers = Invoke-PnPCommandWithThrottling -Command {
                                                        Get-PnPGroupMember -Identity $groupTitle -ErrorAction SilentlyContinue
                                                    } -OperationDescription "Get SharePoint group members for EEEU check: $groupTitle"
                                                    $eeeuInGroup = $groupMembers | Where-Object { $_.LoginName -like $EEEU }
                                                
                                                    if ($eeeuInGroup) {
                                                        Write-DebugInfo "    ✓ Found EEEU in SharePoint group '$groupTitle' with permissions: $($RoleAssignment.RoleDefinitionBindings.Name -join ', ')" -ForegroundColor Green
                                                    
                                                        # Only grant access if user is internal
                                                        if ($user -like "*@$t.onmicrosoft.com" -or $user -like "*@*.onmicrosoft.com") {
                                                            $userFound = $true
                                                            $accessType += "; Everyone except external users (via $groupTitle)"
                                                            $everyoneExceptExternalFound = $true
                                                            Write-Host "✓ Found $user has access via 'Everyone except external users' (via group '$groupTitle') on '$($site.url)'" -ForegroundColor Green
                                                            break
                                                        }
                                                        else {
                                                            Write-Host "    User is external, not granting access via 'Everyone except external users'" -ForegroundColor DarkYellow
                                                        }
                                                    }
                                                }
                                                catch {
                                                    Write-Host "    Could not check group membership for '$groupTitle': $($_.Exception.Message)" -ForegroundColor DarkYellow
                                                }
                                            }
                                        }
                                        catch {
                                            # Skip role assignments that fail to load
                                            Write-Host "    Warning: Could not process role assignment for group check: $($_.Exception.Message)" -ForegroundColor DarkYellow
                                        }
                                    }
                                }
                            
                                if (-not $everyoneExceptExternalFound) {
                                    Write-Host "    ✗ No EEEU with meaningful permissions found (direct or via groups)" -ForegroundColor DarkYellow
                                }
                            }
                            else {
                                Write-Host "    No role assignments found" -ForegroundColor DarkGray
                            }
                        }
                        catch {
                            Write-Host "    Error checking EEEU permissions: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "    EEEU not present in site users - skipping EEEU check" -ForegroundColor DarkGray
                    }
                
                    if ($everyoneExceptExternalFound) {
                        $groupMemberships += "Everyone except external users"
                    }
                    else {
                        Write-Host "Did not find meaningful 'Everyone except external users' access for this user" -ForegroundColor DarkGray
                    }
                }
                catch {
                    Write-Host "Error checking 'Everyone except external users': $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            else {
                Write-Host "EEEU check is disabled - skipping 'Everyone except external users' permissions check" -ForegroundColor DarkGray
            }
            
            if ($userFound) {
                #Collecting Export Properties for CSV File
                $ExportItem = New-Object PSObject
                $ExportItem  | Add-Member -MemberType NoteProperty -name "SiteName" -value $Site.Title
                $ExportItem  | Add-Member -MemberType NoteProperty -name "URL" -value  $site.Url
                $ExportItem  | Add-Member -MemberType NoteProperty -name "User" -value $User
                $ExportItem  | Add-Member -MemberType NoteProperty -name "Owner" -value $site.Owner
                
                # Clean up AccessType by removing any entries containing "Limited Access"
                $cleanAccessType = ($accessType.TrimStart('; ').Trim() -split '; ') | Where-Object { $_ -notlike "*Limited Access*" }
                $finalAccessType = $cleanAccessType -join '; '
                
                $ExportItem  | Add-Member -MemberType NoteProperty -name "AccessType" -value $finalAccessType
                $siteOutput += $ExportItem
                
                Write-LogEntry -LogName:$Log -LogEntryText "Found $user on '$($site.url)' with access: $finalAccessType"
            } 
            else {
                Write-DebugInfo "$user WAS NOT FOUND on '$($site.url)'" -ForegroundColor Magenta
                Write-LogEntry -LogName:$Log -LogEntryText "$user WAS NOT FOUND on '$($site.url)'"
            }
        }
        catch {
            Write-DebugInfo "$user WAS NOT FOUND on '$($site.url)'" -ForegroundColor Magenta
            Write-LogEntry -LogName:$Log -LogEntryText "$user WAS NOT FOUND on '$($site.url)'"
        }
        
        Write-Host ""
        Write-LogEntry -LogName:$Log -LogEntryText ""    
    }
    
    # Write output for this site to CSV file immediately after processing all users for this site
    if ($siteOutput.Count -gt 0) {
        if ($firstWrite) {
            # First write includes headers
            $siteOutput | Export-Csv $outputfile -NoTypeInformation
            $firstWrite = $false
            Write-DebugInfo "Exported $($siteOutput.Count) user(s) for site '$($site.Title)' (with headers)" -ForegroundColor Green
        }
        else {
            # Subsequent writes append without headers
            $siteOutput | Export-Csv $outputfile -NoTypeInformation -Append
            Write-DebugInfo "Exported $($siteOutput.Count) user(s) for site '$($site.Title)' (appended)" -ForegroundColor Green
        }
        
        Write-LogEntry -LogName:$Log -LogEntryText "Exported $($siteOutput.Count) user(s) for site '$($site.Title)' to CSV"
        
        # Clear the site output array to free memory
        $siteOutput = @()
    }
    else {
        Write-DebugInfo "No users with access found for site '$($site.Title)'" -ForegroundColor DarkGray
        Write-LogEntry -LogName:$Log -LogEntryText "No users with access found for site '$($site.Title)'"
    }
    
    Write-StatusMessage "Completed processing site: $($site.Title)" -ForegroundColor Yellow
    Write-DebugInfo "----------------------------------------" -ForegroundColor DarkGray
    Write-LogEntry -LogName:$Log -LogEntryText "Completed processing site: $($site.Title)"
    Write-LogEntry -LogName:$Log -LogEntryText "----------------------------------------"
}

#Output Results
Write-Host ""
Write-StatusMessage "All sites processed successfully!" -ForegroundColor Green
Write-StatusMessage "Output file saved to $outputfile" -ForegroundColor Green
Write-Host ""
Write-StatusMessage "Log file saved to $log" -ForegroundColor Green

# Disconnect from SharePoint Online
Disconnect-PnPOnline
