<#
.SYNOPSIS
    Collects comprehensive information about SharePoint Online sites and users in a tenant.

.DESCRIPTION
    This script connects to a SharePoint Online tenant and collects detailed information about sites and 
    their users, including site collection properties, permissions, groups, users, and configuration settings.
    The information is exported to a CSV file for analysis.

.PARAMETER tenantname
    Your SharePoint Online tenant name (without .sharepoint.com)

.PARAMETER appID
    The Microsoft Entra (Azure AD) application ID for authentication

.PARAMETER thumbprint
    The certificate thumbprint for app-based authentication

.PARAMETER tenant
    Your tenant ID (GUID)

.PARAMETER Debug
    When set to $True, includes debug-level messages in the log file

.PARAMETER inputfile
    Optional. Path to a CSV file containing a list of sites to process. If not provided, all sites will be processed.
    CSV should have a header of "URL" with site URLs in the first column.

.PARAMETER maxcount
    Optional. The maximum number of users or groups to list in a single cell before showing a summary. Default is 150.

.PARAMETER ExpandDynamicDLs
    Optional. Specifies whether to expand dynamic Microsoft 365 groups to list their members.
    Set to $false to list the membership rule instead of members for dynamic groups. Default is $false.

.NOTES
    File Name      : Get-SPSitesAndUsersInfo.ps1
    Author         : Mike Lee / Andrew Thompson
    Prerequisite   : PnP.PowerShell module installed
    Date           : 6/19/25     
    Version        : 2.1

    Requirements:
        - PnP.PowerShell module installed (Tested with PNP 2.12.0)
        - PowerShell 7.4 or higer
        - Appropriate permissions granted to the Azure AD application
            - Microsoft Graph| Application | Directory.Read.All
            - SharePoint |Application | Sites.FullControl.All
        - Certificate-based authentication configured

    The script collects information including:
    - Site properties (template, sharing capabilities, storage, etc.)
    - Site collection administrators
    - SharePoint groups and their roles
    - Group members and direct site users
    - Associated Microsoft 365 Groups and their members
    - EEEU (Everyone except external users) presence and permissions
    - Community site settings
    - Version policy settings
    - Sharing link presence


.EXAMPLE
    # To process only specific sites:
    $tenantname = "contoso"
    $appID = "12345678-1234-1234-1234-1234567890ab"
    $thumbprint = "A1B2C3D4E5F6G7H8I9J0K1L2M3N4O5P6Q7R8S9T0"
    $tenant = "87654321-4321-4321-4321-ba0987654321"
    $inputfile = "C:\temp\sitelist-contoso.csv"
    $maxcount = 150
    .\Get-SPSitesAndUsersInfo.ps1
#>

# Set Variables
$tenantname = "m365x61250205" #This is your tenant name
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"  #This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9" #This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51" #This is your Tenant ID
$Debug = $false # Set to $True to include debug-level messages in the log file, $False to exclude them
$maxcount = 150 # The maximum number of users or groups to list in a cell. If exceeded, a summary message is shown to avoid Excel cell limits.
$ExpandDynamicDLs = $false # Default to expanding dynamic groups. Set to $false to store rule instead of members for dynamic groups.

#Initialize Parameters - Do not change
$sites = @() # Array to hold site objects to be processed
$inputfile = $null # Path to the optional input CSV file for specific sites
$outputfile = $null # Path for the output CSV file
$log = $null # Path for the log file
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss" # Current date and time for unique file naming
$maxRetries = 5  # Maximum number of retry attempts for PnP cmdlets
$initialRetryDelay = 5  # Initial retry delay in seconds for PnP cmdlets

#Input / Output and Log Files
#$inputfile = "C:\temp\sitelist-m365x61250205.csv" # Example: This is the input file with list of sites to process. If not provided, all sites will be processed.
$outputfile = "$env:TEMP\" + 'SPSites_and_Users_Info_' + $date + '_' + "output.csv" # Define output CSV file path
$log = "$env:TEMP\" + 'SPSites_and_Users_Info_' + $date + '_' + "logfile.log" # Define log file path

# OPTIMIZATION: Initialize AAD User Cache
$aadUserCache = @{} # Hashtable to cache Azure AD user objects to reduce API calls
$aadUserNotFoundMarker = [PSCustomObject]@{ NotFound = $true } # Marker for users not found in AAD to avoid repeated lookups

#This is the logging function
Function Write-LogEntry {
    param(
        [string] $LogName, # Path to the log file
        [string] $LogEntryText, # Text to write to the log
        [string] $LogLevel = "INFO"  # Default log level is INFO (INFO, DEBUG, WARNING, ERROR)
    )
    if ($LogName -ne $null) {
        # Skip DEBUG level messages if Debug mode is set to False
        if ($LogLevel -eq "DEBUG" -and $Debug -eq $False) {
            return
        }
        
        # log the date and time in the text file along with the data passed
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : [$LogLevel] $LogEntryText" | Out-File -FilePath $LogName -append;
    }
}

# Function to handle throttling with exponential backoff for PnP cmdlets
Function Invoke-PnPWithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock] $ScriptBlock, # The PnP command to execute
        
        [Parameter(Mandatory = $false)]
        [string] $Operation = "PnP Operation", # Description of the operation for logging
        
        [Parameter(Mandatory = $false)]
        [int] $MaxRetries = 5, # Maximum number of retries for this specific operation
        
        [Parameter(Mandatory = $false)]
        [int] $InitialRetryDelay = 5, # Initial delay in seconds before retrying
        
        [Parameter(Mandatory = $false)]
        [string] $LogName # Path to the log file
    )
    
    $retryCount = 0
    $success = $false
    $result = $null
    $retryDelay = $InitialRetryDelay
    
    do {
        try {
            # Execute the provided script block
            $result = & $ScriptBlock
            $success = $true
            return $result
        }
        catch {
            $exceptionDetails = $_.Exception.ToString()
            
            # Check for common throttling-related HTTP status codes or messages
            if (($exceptionDetails -like "*429*") -or 
                ($exceptionDetails -like "*throttl*") -or 
                ($exceptionDetails -like "*too many requests*") -or
                ($exceptionDetails -like "*request limit exceeded*")) {
                
                $retryCount++
                
                # Check if maximum retries have been reached
                if ($retryCount -ge $MaxRetries) {
                    Write-LogEntry -LogName $LogName -LogEntryText "Max retries ($MaxRetries) reached for $Operation. Giving up." -LogLevel "ERROR" 
                    throw $_ # Re-throw the original exception
                }
                
                # Parse Retry-After header from the exception response if available
                $retryAfterValue = $null
                if ($_.Exception.Response -and $_.Exception.Response.Headers -and $_.Exception.Response.Headers["Retry-After"]) {
                    $retryAfterValue = [int]$_.Exception.Response.Headers["Retry-After"]
                    $retryDelay = $retryAfterValue # Use server-suggested delay
                    Write-LogEntry -LogName $LogName -LogEntryText "Throttling detected for $Operation. Server requested retry after $retryAfterValue seconds." -LogLevel "WARNING"
                }
                else {
                    # Use exponential backoff if no Retry-After header is present
                    $retryDelay = [Math]::Min(60, $retryDelay * 2) # Double the delay, max 60 seconds
                    Write-LogEntry -LogName $LogName -LogEntryText "Throttling detected for $Operation. Using exponential backoff: waiting $retryDelay seconds before retry $retryCount of $MaxRetries." -LogLevel "WARNING"
                }
                
                Write-Host "Throttling detected for $Operation. Waiting $retryDelay seconds before retry $retryCount of $MaxRetries." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryDelay # Wait before retrying
            }
            else {
                # If not a throttling error, re-throw the original exception
                throw $_
            }
        }
    } while (-not $success -and $retryCount -lt $MaxRetries)
}

# OPTIMIZATION: Function to get AAD User from Cache or API
Function Get-CachedPnPAzureADUser {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Identity, # UserPrincipalName or ObjectId of the AAD user
        [string] $LogName # Path to the log file
    )

    $cacheKey = $Identity.ToLower() # Use lowercase for consistent cache key
    # Check if user is already in cache
    if ($aadUserCache.ContainsKey($cacheKey)) {
        $cachedUser = $aadUserCache[$cacheKey]
        # Check if the cached entry is the 'not found' marker
        if ($cachedUser -eq $aadUserNotFoundMarker) {
            Write-LogEntry -LogName $LogName -LogEntryText "AAD User '$Identity' previously not found (from cache)." -LogLevel "DEBUG"
            return $null # User was previously confirmed as not found
        }
        Write-LogEntry -LogName $LogName -LogEntryText "AAD User '$Identity' found in cache." -LogLevel "DEBUG"
        return $cachedUser # Return cached user object
    }
    try {
        # User not in cache, fetch from API
        Write-LogEntry -LogName $LogName -LogEntryText "Fetching AAD User '$Identity' from API." -LogLevel "DEBUG"
        $upn = $Identity.Split('|')[-1] # Extract UPN
        $user = Invoke-PnPWithRetry -ScriptBlock { 
            Get-PnPAzureADUser -Identity $upn -ErrorAction SilentlyContinue # Suppress non-terminating errors for checking existence
        } -Operation "Get-PnPAzureADUser for $upn (Cached)" -LogName $LogName
        
        if ($user) {
            $aadUserCache[$cacheKey] = $user # Add found user to cache
            return $user
        }
        else {
            Write-LogEntry -LogName $LogName -LogEntryText "AAD User '$Identity' not found via API." -LogLevel "DEBUG"
            $aadUserCache[$cacheKey] = $aadUserNotFoundMarker # Mark as not found in cache
            return $null
        }
    }
    catch {
        # Handle errors during API call
        Write-LogEntry -LogName $LogName -LogEntryText "Error fetching AAD User '$Identity': $_.Exception.Message. Marking as not found." -LogLevel "WARNING"
        $aadUserCache[$cacheKey] = $aadUserNotFoundMarker # Mark as not found in cache on error
        return $null
    }
}

# Helper function to determine if EEEU (Everyone Except External Users) has meaningful permissions on a site
Function Set-EEEUPresentIfApplicable {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrlToUpdate, # The URL of the site being checked
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[string]]$Roles, # List of roles assigned to EEEU
        [Parameter(Mandatory = $true)]
        [string]$ContextForLog # Context of how EEEU was found (e.g., direct assignment, group membership)
    )
    
    # If no roles are provided, EEEU doesn't have permissions through this context
    if ($null -eq $Roles -or $Roles.Count -eq 0) {
        Write-LogEntry -LogName $Log -LogEntryText "Set-EEEUPresentIfApplicable: No roles provided for EEEU via $ContextForLog on site $SiteUrlToUpdate. No change to 'EEEU Count'." -LogLevel "DEBUG"
        return
    }

    $hasOnlyLimitedAccess = $true # Flag to check if EEEU only has "Limited Access"
    $hasAnyPermission = $false # Flag to check if EEEU has any permission at all
    
    # Iterate through the roles to determine if any are beyond "Limited Access"
    foreach ($roleName in $Roles) {
        $hasAnyPermission = $true
        if ($roleName -ne "Limited Access") {
            $hasOnlyLimitedAccess = $false # Found a role other than "Limited Access"
            break
        }
    }

    # If EEEU has any permission and it's not just "Limited Access", increment 'EEEU Count'
    if ($hasAnyPermission -and -not $hasOnlyLimitedAccess) {
        Write-LogEntry -LogName $Log -LogEntryText "EEEU has meaningful permissions ($($Roles -join ',')) via $ContextForLog. Incrementing 'EEEU Count' for site $SiteUrlToUpdate" -LogLevel "INFO"
        if ($siteCollectionData.ContainsKey($SiteUrlToUpdate)) {
            # Increment the EEEU count
            $siteCollectionData[$SiteUrlToUpdate]["EEEU Count"]++
        }
        else {
            Write-LogEntry -LogName $Log -LogEntryText "Set-EEEUPresentIfApplicable: Site $SiteUrlToUpdate not found in siteCollectionData. Cannot increment 'EEEU Count'." -LogLevel "WARNING"
        }
    }
    else {
        # EEEU has no roles or only "Limited Access"
        Write-LogEntry -LogName $Log -LogEntryText "Set-EEEUPresentIfApplicable: EEEU via $ContextForLog on site $SiteUrlToUpdate has no roles or only 'Limited Access' ($($Roles -join ',')). 'EEEU Count' remains unchanged (currently $($siteCollectionData[$SiteUrlToUpdate]['EEEU Count']))." -LogLevel "DEBUG"
    }
}


# Define the connection parameters for reuse across PnP cmdlets
$connectionParams = @{
    ClientId      = $appID         # Azure AD App ID for authentication
    Thumbprint    = $thumbprint    # Certificate thumbprint for app-based authentication
    Tenant        = $tenant         # Tenant ID (GUID)
    WarningAction = 'SilentlyContinue' # Suppress PnP warnings that are not errors
}

#Connect to SharePoint Admin Center initially
try {
    $adminUrl = 'https://' + $tenantname + '-admin.sharepoint.com' # Construct Admin Center URL
    
    # Connect using retry logic
    Invoke-PnPWithRetry -ScriptBlock { 
        Connect-PnPOnline -Url $adminUrl @connectionParams 
    } -Operation "Connect to SharePoint Admin Center" -LogName $Log
    
    Write-LogEntry -LogName $Log -LogEntryText "Successfully connected to SharePoint Admin Center: $adminUrl"
}
catch {
    # Handle connection failure
    Write-Host "Error connecting to SharePoint Admin Center ($adminUrl): $_" -ForegroundColor Red
    Write-LogEntry -LogName $Log -LogEntryText "Error connecting to SharePoint Admin Center ($adminUrl): $_" -LogLevel "ERROR"
    exit # Exit script if initial connection fails
}

# Get Site List: either from an input file or by querying all tenant sites
if ($inputfile -and (Test-Path -Path $inputfile)) {
    # Input file provided and exists
    try {
        $sites = Import-csv -path $inputfile -Header 'URL' # Import site URLs from CSV
        Write-LogEntry -LogName $Log -LogEntryText "Using sites from input file: $inputfile"
        Write-Host "Reading sites from input file: $inputfile" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Error reading input file '$inputfile': $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error reading input file '$inputfile': $_" -LogLevel "ERROR"
        exit # Exit if input file reading fails
    }
}
else {
    # No input file, or file not found; get all sites from the tenant
    Write-Host "Getting site list from tenant (this might take a while)..." -ForegroundColor Yellow
    Write-LogEntry -LogName $Log -LogEntryText "Getting sites using Get-PnPTenantSite (no input file specified or found)"
    try {
        # Ensure connection to Admin Center before getting tenant sites
        Invoke-PnPWithRetry -ScriptBlock { 
            Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop 
        } -Operation "Connect to SharePoint Admin Center (before Get-PnPTenantSite)" -LogName $Log
        
        # Retrieve all tenant sites, excluding MySites and RedirectSites
        $sites = Invoke-PnPWithRetry -ScriptBlock { 
            Get-PnPTenantSite | Where-Object { $_.Url -notlike "*-my.sharepoint.com*" -and $_.Template -ne "RedirectSite#0" }
        } -Operation "Get-PnPTenantSite" -LogName $Log
        
        Write-Host "Found $($sites.Count) sites." -ForegroundColor Green
        Write-LogEntry -LogName $Log -LogEntryText "Retrieved $($sites.Count) sites using Get-PnPTenantSite."
    }
    catch {
        Write-Host "Error getting site list from tenant: $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error getting site list from tenant: $_" -LogLevel "ERROR"
        exit # Exit if fetching all sites fails
    }
}

$siteCollectionData = @{} # Hashtable to store detailed data for each site collection before exporting

# Function to update the in-memory data store for a site collection
Function Update-SiteCollectionData {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [object] $SiteProperties, 
        [string] $SPGroupName = "",
        [string] $SPGroupRoles = "", 
        [string] $AssociatedSPGroup = "", 
        [string] $SPUserName = "",
        [string] $SPUserTitle = "",
        [string] $SPUserEmail = "",
        [string] $SPUserLoginName = "", 
        [string] $EntraGroupOwner = "",
        [string] $EntraGroupOwnerEmail = "",
        [string] $EntraGroupMember = "",
        [string] $EntraGroupMemberEmail = "",
        [object] $AADGroups = $null,
        [string] $SiteAdminName = "",
        [string] $SiteAdminEmail = "",
        [string] $DefaultTrimMode = "",
        [int] $DefaultExpireAfterDays = -1,
        [int] $MajorVersionLimit = -1,
        [bool] $IsCommunity = $false,
        [bool] $AllowMembersEditMembership = $false,
        [bool] $MembersCanShare = $false,
        [bool] $ContainsSubSites = $false,
        [string] $EntraGroupMembershipRule = "" # New parameter for dynamic group rule
    )

    # Initialize the site's data structure if it doesn't exist
    if (-not $siteCollectionData.ContainsKey($SiteUrl)) {
        $siteCollectionData[$SiteUrl] = @{
            "URL"                            = $SiteProperties.Url
            "Owner"                          = $SiteProperties.Owner
            "IB Mode"                        = ($SiteProperties.InformationBarrierMode -join ',')
            "IB Segment"                     = ($SiteProperties.InformationBarrierSegments -join ',')
            "Group ID"                       = $SiteProperties.GroupId
            "RelatedGroupId"                 = $SiteProperties.RelatedGroupId
            "IsHubSite"                      = $SiteProperties.IsHubSite
            "Template"                       = $SiteProperties.Template
            "SiteDefinedSharingCapability"   = $SiteProperties.SiteDefinedSharingCapability
            "SharingCapability"              = $SiteProperties.SharingCapability
            "DisableCompanyWideSharingLinks" = $SiteProperties.DisableCompanyWideSharingLinks
            "Custom Script Allowed"          = if ($SiteProperties.DenyAddAndCustomizePages -eq "Enabled") { $false } else { $true }
            "IsTeamsConnected"               = $SiteProperties.IsTeamsConnected
            "IsTeamsChannelConnected"        = $SiteProperties.IsTeamsChannelConnected
            "TeamsChannelType"               = $SiteProperties.TeamsChannelType
            "StorageQuota"                   = if ($SiteProperties.StorageQuota) { $SiteProperties.StorageQuota } else { 0 }
            "StorageUsageCurrent"            = if ($SiteProperties.StorageUsageCurrent) { $SiteProperties.StorageUsageCurrent } else { 0 }
            "LockState"                      = $SiteProperties.LockState
            "LastContentModifiedDate"        = $SiteProperties.LastContentModifiedDate
            "ArchiveState"                   = $SiteProperties.ArchiveState
            "DefaultTrimMode"                = $DefaultTrimMode
            "DefaultExpireAfterDays"         = $DefaultExpireAfterDays
            "MajorVersionLimit"              = $MajorVersionLimit
            "Community Site"                 = $IsCommunity
            "AllowMembersEditMembership"     = $AllowMembersEditMembership
            "MembersCanShare"                = $MembersCanShare
            "Contains SubSites"              = $ContainsSubSites
            "SP Groups On Site"              = [System.Collections.Generic.List[string]]::new()
            "SP Group Roles Per Group"       = [System.Collections.Generic.Dictionary[string, string]]::new() 
            "SP Users"                       = [System.Collections.Generic.List[PSObject]]::new() 
            "Entra Group Owners"             = [System.Collections.Generic.List[PSObject]]::new() 
            "Entra Group Members"            = [System.Collections.Generic.List[PSObject]]::new() 
            "Entra Group Membership Rule"    = "" # New field for dynamic group rule
            "Entra Group Details"            = $null
            "Site Collection Admins"         = [System.Collections.Generic.List[PSObject]]::new() 
            "Site Level Users"               = [System.Collections.Generic.List[PSObject]]::new() 
            "Sharing Links Count"            = 0  # Count of sharing links found on the site
            "EEEU Count"                     = 0  # Count of EEEU present on the site
        }
    }
    
    # Update specific site properties if they are passed as parameters
    if ($PSBoundParameters.ContainsKey('IsCommunity')) { $siteCollectionData[$SiteUrl]["Community Site"] = $IsCommunity }
    if ($PSBoundParameters.ContainsKey('ContainsSubSites')) { $siteCollectionData[$SiteUrl]["Contains SubSites"] = $ContainsSubSites }
    if (-not [string]::IsNullOrEmpty($DefaultTrimMode)) { $siteCollectionData[$SiteUrl]["DefaultTrimMode"] = $DefaultTrimMode }
    if ($DefaultExpireAfterDays -ge 0) { $siteCollectionData[$SiteUrl]["DefaultExpireAfterDays"] = $DefaultExpireAfterDays } # -1 indicates not set
    if ($MajorVersionLimit -ge 0) { $siteCollectionData[$SiteUrl]["MajorVersionLimit"] = $MajorVersionLimit } # -1 indicates not set
    if ($PSBoundParameters.ContainsKey('AllowMembersEditMembership')) { $siteCollectionData[$SiteUrl]["AllowMembersEditMembership"] = $AllowMembersEditMembership }
    if ($PSBoundParameters.ContainsKey('MembersCanShare')) { $siteCollectionData[$SiteUrl]["MembersCanShare"] = $MembersCanShare }
    if ($PSBoundParameters.ContainsKey('EntraGroupMembershipRule') -and -not [string]::IsNullOrWhiteSpace($EntraGroupMembershipRule)) { $siteCollectionData[$SiteUrl]["Entra Group Membership Rule"] = $EntraGroupMembershipRule }

    # Check if the SPGroupName indicates the presence of Sharing Links
    if (-not [string]::IsNullOrWhiteSpace($SPGroupName) -and $SPGroupName -like "SharingLinks*") {
        $siteCollectionData[$SiteUrl]["Sharing Links Count"]++
    }

    # Add Site Collection Administrator if provided and not already present
    if (-not [string]::IsNullOrWhiteSpace($SiteAdminEmail)) {
        $checkEmail = $SiteAdminEmail.ToLower() 
        # Ensure admin is not added multiple times by checking email
        if (-not ($siteCollectionData[$SiteUrl]["Site Collection Admins"].Email -contains $checkEmail)) {
            $adminObject = [PSCustomObject]@{ Name = $SiteAdminName; Email = $SiteAdminEmail }
            $siteCollectionData[$SiteUrl]["Site Collection Admins"].Add($adminObject)
        }
    }

    # Store Entra (Azure AD) Group details if provided
    if ($AADGroups) {
        $siteCollectionData[$SiteUrl]["Entra Group Details"] = [PSCustomObject]@{
            DisplayName = $AADGroups.DisplayName; Alias = $AADGroups.MailNickname
            AccessType = $AADGroups.Visibility; WhenCreated = $AADGroups.CreatedDateTime
        }
    }

    # Process SharePoint Group information
    if (-not [string]::IsNullOrWhiteSpace($SPGroupName)) {
        # Add SP Group to the list if not already present
        if (-not $siteCollectionData[$SiteUrl]["SP Groups On Site"].Contains($SPGroupName)) {
            $siteCollectionData[$SiteUrl]["SP Groups On Site"].Add($SPGroupName)
        }
        # Add or update SP Group roles
        if (-not [string]::IsNullOrWhiteSpace($SPGroupRoles)) {
            if (-not $siteCollectionData[$SiteUrl]["SP Group Roles Per Group"].ContainsKey($SPGroupName)) {
                # Add new group and its roles
                $siteCollectionData[$SiteUrl]["SP Group Roles Per Group"].Add($SPGroupName, $SPGroupRoles)
            }
            else { 
                # If group exists, merge new roles with existing ones, ensuring uniqueness
                if ($siteCollectionData[$SiteUrl]["SP Group Roles Per Group"][$SPGroupName] -ne $SPGroupRoles) {
                    $existingRoles = $siteCollectionData[$SiteUrl]["SP Group Roles Per Group"][$SPGroupName].Split(',') | Where-Object { $_ }
                    $newRoles = $SPGroupRoles.Split(',') | Where-Object { $_ }
                    $allRoles = ($existingRoles + $newRoles | Select-Object -Unique) -join ','
                    $siteCollectionData[$SiteUrl]["SP Group Roles Per Group"][$SPGroupName] = $allRoles
                }
            }
        }
    }

    # Add SharePoint User associated with an SP Group
    if (-not [string]::IsNullOrWhiteSpace($SPUserName)) {
        $userObject = [PSCustomObject]@{
            AssociatedSPGroup = $AssociatedSPGroup 
            Name              = $SPUserName
            Title             = $SPUserTitle
            Email             = $SPUserEmail
        }
        $siteCollectionData[$SiteUrl]["SP Users"].Add($userObject)
    }

    # Add Entra Group Owner if provided and not already present
    if (-not [string]::IsNullOrWhiteSpace($EntraGroupOwnerEmail)) {
        $checkEmail = $EntraGroupOwnerEmail.ToLower()
        # Ensure owner is not added multiple times by checking email
        if (-not ($siteCollectionData[$SiteUrl]["Entra Group Owners"].Email -contains $checkEmail)) {
            $ownerObject = [PSCustomObject]@{ Name = $EntraGroupOwner; Email = $EntraGroupOwnerEmail }
            $siteCollectionData[$SiteUrl]["Entra Group Owners"].Add($ownerObject)
        }
    }

    # Add Entra Group Member if provided and not already present
    if (-not [string]::IsNullOrWhiteSpace($EntraGroupMemberEmail)) {
        $checkEmail = $EntraGroupMemberEmail.ToLower()
        # Ensure member is not added multiple times by checking email
        if (-not ($siteCollectionData[$SiteUrl]["Entra Group Members"].Email -contains $checkEmail)) {
            $memberObject = [PSCustomObject]@{ Name = $EntraGroupMember; Email = $EntraGroupMemberEmail }
            $siteCollectionData[$SiteUrl]["Entra Group Members"].Add($memberObject)
        }
    }
}

$totalSites = $sites.Count # Total number of sites to process
$processedCount = 0 # Counter for processed sites

# Define CSV Headers for the output file
$csvHeaders = "URL,Owner,IB Mode,IB Segment,Group ID,RelatedGroupId,IsHubSite,Template,SiteDefinedSharingCapability," + 
"SharingCapability,DisableCompanyWideSharingLinks,AllowMembersEditMembership,MembersCanShare,Custom Script Allowed,IsTeamsConnected,IsTeamsChannelConnected," + 
"TeamsChannelType,StorageQuota (MB),StorageUsageCurrent (MB),LockState,LastContentModifiedDate,ArchiveState," + 
"DefaultTrimMode,DefaultExpireAfterDays,MajorVersionLimit, Entra Group Alias," + 
"Entra Group AccessType,Entra Group WhenCreated,Sharing Links Count," + 
"EEEU Count,Community Site,Contains SubSites,SP Groups On Site,SP Groups Roles,Site Collection Admins (Name <Email>)," +
"Site Level Users (Name <Email> [Roles]), SP Users (Group: Name <Email>),Entra Group Owners (Name <Email>),Entra Group Members (Name <Email>)"

# Create the output CSV file and write the headers
Set-Content -Path $outputfile -Value $csvHeaders -Encoding UTF8
Write-Host "Created output file with headers: $outputfile" -ForegroundColor Green
Write-LogEntry -LogName $Log -LogEntryText "Created output file with headers: $outputfile"

# Function to format and export data for a single site collection to the CSV file
function Export-SiteCollectionToCSV {
    param(
        [string] $SiteUrl, # URL of the site to export
        [string] $CsvPath,  # Path to the CSV output file
        [int] $MaxCountForCell # Max items before summarizing
    )
    
    $siteData = $siteCollectionData[$SiteUrl] # Retrieve the collected data for the site
    if (-not $siteData) {
        Write-LogEntry -LogName $Log -LogEntryText "Error: No data found for site $SiteUrl when attempting to export" -LogLevel "ERROR"
        return
    }

    # --- Format complex data types with maxcount check ---

    # SP Users
    $spUsersList = $siteData."SP Users"
    $spUsersCount = $spUsersList.Count
    $spUsersFormatted = ""
    if ($spUsersCount -gt $MaxCountForCell) {
        $spUsersFormatted = "max limit reached: $spUsersCount users found"
    }
    else {
        $spUsersFormatted = ($spUsersList | ForEach-Object {
                $emailStr = if ($_.Email) { $_.Email } else { "N/A" }
                "$($_.AssociatedSPGroup):$($_.Name) <$emailStr>"
            }) -join ';'
    }
  
    # Entra Group Owners
    $entraOwnersList = $siteData."Entra Group Owners"
    $entraOwnersCount = $entraOwnersList.Count
    $entraOwnersFormatted = ""
    if ($entraOwnersCount -gt $MaxCountForCell) {
        $entraOwnersFormatted = "max limit reached: $entraOwnersCount owners found"
    }
    else {
        $entraOwnersFormatted = ($entraOwnersList | ForEach-Object {
                $emailStr = if ($_.Email) { $_.Email } else { "N/A" }
                "$($_.Name) <$emailStr>"
            }) -join ';'
    }

    # Entra Group Members
    $entraMembersList = $siteData."Entra Group Members"
    $entraMembersCount = $entraMembersList.Count
    $entraMembersFormatted = ""
    if ($entraMembersCount -eq 0 -and -not [string]::IsNullOrWhiteSpace($siteData."Entra Group Membership Rule")) {
        # If no members are listed but a rule exists (dynamic group, not expanded)
        $entraMembersFormatted = "Dynamic Group Rule: $($siteData."Entra Group Membership Rule")"
    }
    elseif ($entraMembersCount -gt $MaxCountForCell) {
        $entraMembersFormatted = "max limit reached: $entraMembersCount members found"
    }
    else {
        $entraMembersFormatted = ($entraMembersList | ForEach-Object {
                $emailStr = if ($_.Email) { $_.Email } else { "N/A" }
                "$($_.Name) <$emailStr>"
            }) -join ';'
    }
    
    # Site Collection Admins
    $siteAdminsList = $siteData."Site Collection Admins"
    $siteAdminsCount = $siteAdminsList.Count
    $siteAdminsFormatted = ""
    if ($siteAdminsCount -gt $MaxCountForCell) {
        $siteAdminsFormatted = "max limit reached: $siteAdminsCount admins found"
    }
    else {
        $siteAdminsFormatted = ($siteAdminsList | ForEach-Object {
                $emailStr = if ($_.Email) { $_.Email } else { "N/A" }
                "$($_.Name) <$emailStr>"
            }) -join ';'
    }
    
    # Site Level Users
    $siteLevelUsersList = $siteData."Site Level Users"
    $siteLevelUsersCount = $siteLevelUsersList.Count
    $siteLevelUsersFormatted = ""
    if ($siteLevelUsersCount -gt $MaxCountForCell) {
        $siteLevelUsersFormatted = "max limit reached: $siteLevelUsersCount users/groups found"
    }
    else {
        $siteLevelUsersFormatted = ($siteLevelUsersList | ForEach-Object {
                $emailStr = if ($_.Email) { $_.Email } else { "N/A" }
                "$($_.Name) <$emailStr> [$($_.Roles)]" 
            }) -join ';'
    }
    
    # SP Group Roles
    $spGroupRolesDict = $siteData."SP Group Roles Per Group"
    $spGroupRolesCount = $spGroupRolesDict.Keys.Count
    $spGroupRolesFormatted = ""
    if ($spGroupRolesCount -gt $MaxCountForCell) {
        $spGroupRolesFormatted = "max limit reached: $spGroupRolesCount groups with roles found"
    }
    else {
        $spGroupRolesFormatted = ($spGroupRolesDict.GetEnumerator() | ForEach-Object {
                "$($_.Key): [$($_.Value)]" # Format as "GroupName: [Role1,Role2]"
            }) -join ';'
    }
    
    # SP Groups On Site
    $spGroupsList = $siteData."SP Groups On Site"
    $spGroupsCount = $spGroupsList.Count
    $spGroupsOnSiteFormatted = ""
    if ($spGroupsCount -gt $MaxCountForCell) {
        $spGroupsOnSiteFormatted = "max limit reached: $spGroupsCount groups found"
    }
    else {
        $spGroupsOnSiteFormatted = ($spGroupsList -join ';')
    }

    # Create a PSCustomObject with all properties matching the CSV headers
    $exportItem = [PSCustomObject]@{
        URL                                       = $siteData.URL
        Owner                                     = $siteData.Owner
        "IB Mode"                                 = $siteData."IB Mode"
        "IB Segment"                              = $siteData."IB Segment"
        "Group ID"                                = $siteData."Group ID"
        RelatedGroupId                            = $siteData.RelatedGroupId
        IsHubSite                                 = $siteData.IsHubSite
        Template                                  = $siteData.Template
        SiteDefinedSharingCapability              = $siteData.SiteDefinedSharingCapability
        SharingCapability                         = $siteData.SharingCapability
        DisableCompanyWideSharingLinks            = $siteData.DisableCompanyWideSharingLinks
        "Custom Script Allowed"                   = if ($siteData."Custom Script Allowed") { "True" } else { "False" }
        IsTeamsConnected                          = $siteData.IsTeamsConnected
        IsTeamsChannelConnected                   = $siteData.IsTeamsChannelConnected
        TeamsChannelType                          = $siteData.TeamsChannelType
        "StorageQuota (MB)"                       = $siteData.StorageQuota
        "StorageUsageCurrent (MB)"                = $siteData.StorageUsageCurrent
        LockState                                 = $siteData.LockState
        LastContentModifiedDate                   = $siteData.LastContentModifiedDate
        ArchiveState                              = $siteData.ArchiveState
        DefaultTrimMode                           = $siteData.DefaultTrimMode
        DefaultExpireAfterDays                    = if ($siteData.DefaultExpireAfterDays -eq -1) { "NotSet" } else { $siteData.DefaultExpireAfterDays }
        MajorVersionLimit                         = if ($siteData.MajorVersionLimit -eq -1) { "NotSet" } else { $siteData.MajorVersionLimit }
        "Entra Group Alias"                       = if ($siteData."Entra Group Details") { $siteData."Entra Group Details".Alias } else { $null }
        "Entra Group AccessType"                  = if ($siteData."Entra Group Details") { $siteData."Entra Group Details".AccessType } else { $null }
        "Entra Group WhenCreated"                 = if ($siteData."Entra Group Details") { $siteData."Entra Group Details".WhenCreated } else { $null }
        "Sharing Links Count"                     = $siteData."Sharing Links Count"
        "EEEU Count"                              = $siteData."EEEU Count"
        "Community Site"                          = if ($siteData."Community Site") { "True" } else { "False" }
        "Contains SubSites"                       = if ($siteData."Contains SubSites") { "True" } else { "False" }
        "AllowMembersEditMembership"              = if ($siteData."AllowMembersEditMembership") { "True" } else { "False" }
        "MembersCanShare"                         = if ($siteData."MembersCanShare") { "True" } else { "False" }
        "SP Groups On Site"                       = $spGroupsOnSiteFormatted
        "SP Groups Roles"                         = $spGroupRolesFormatted 
        "Site Collection Admins (Name <Email>)"   = $siteAdminsFormatted
        "SP Users (Group: Name <Email>)"          = $spUsersFormatted       
        "Site Level Users (Name <Email> [Roles])" = $siteLevelUsersFormatted 
        "Entra Group Owners (Name <Email>)"       = $entraOwnersFormatted   
        "Entra Group Members (Name <Email>)"      = $entraMembersFormatted  
    }

    try {
        # Append the formatted site data to the CSV file
        $exportItem | Export-Csv -Path $CsvPath -NoTypeInformation -Append -Encoding UTF8
        Write-LogEntry -LogName $Log -LogEntryText "Successfully wrote data for site $SiteUrl to CSV" -LogLevel "DEBUG"
        # Remove the site's data from memory after exporting to free up resources
        $siteCollectionData.Remove($SiteUrl)
    }
    catch {
        Write-Host "Error writing site data ($SiteUrl) to CSV '$CsvPath': $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error writing site data ($SiteUrl) to CSV '$CsvPath': $_" -LogLevel "ERROR"
    }
}

# Main processing loop: Iterate through each site
foreach ($site in $sites) {
    $processedCount++
    $siteUrl = $site.Url 
    Write-Host "Processing site $processedCount/$totalSites : $siteUrl" -ForegroundColor Cyan
    Write-LogEntry -LogName $Log -LogEntryText "Processing site $processedCount/$totalSites : $siteUrl"

    # Initialize variables for the current site
    $siteprops = $null # To store tenant-level site properties
    $currentPnPConnection = $null # To store the PnP connection object for the specific site
    $containsSubSites = $false # Flag for subsite presence
    $webForSiteLevelUsers = $null # To store the PnPWeb object for site-level user processing

    try {
        # Connect to Admin URL to get tenant-level properties for the site
        Invoke-PnPWithRetry -ScriptBlock { 
            Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop 
        } -Operation "Connect to Admin URL for site props $siteUrl" -LogName $Log
        
        # Get tenant-level site properties
        $siteprops = Invoke-PnPWithRetry -ScriptBlock { 
            Get-PnPTenantSite -Identity $siteUrl | Select-Object Url, Owner, InformationBarrierMode, InformationBarrierSegments, GroupId, RelatedGroupId, IsHubSite, Template, SiteDefinedSharingCapability, SharingCapability, DisableCompanyWideSharingLinks, DenyAddAndCustomizePages, IsTeamsConnected, IsTeamsChannelConnected, TeamsChannelType, StorageQuota, StorageUsageCurrent, LockState, LastContentModifiedDate, ArchiveState
        } -Operation "Get-PnPTenantSite for $siteUrl" -LogName $Log

        # If site properties couldn't be retrieved, log error and skip to the next site
        if ($null -eq $siteprops) { Write-LogEntry -LogName $Log -LogEntryText "Failed to retrieve properties for site $siteUrl. Skipping." -LogLevel "ERROR"; continue }

        # Initialize or update the site's data in the main data store
        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops 

        # Nested try-catch for operations requiring connection to the specific site URL
        try {
            Write-LogEntry -LogName $Log -LogEntryText "Connecting to specific site: $siteUrl" -LogLevel "DEBUG"
            # Connect to the specific site
            $currentPnPConnection = Invoke-PnPWithRetry -ScriptBlock { 
                Connect-PnPOnline -Url $siteUrl @connectionParams -ErrorAction Stop 
            } -Operation "Connect to site $siteUrl" -LogName $Log
            Write-LogEntry -LogName $Log -LogEntryText "Successfully connected to specific site: $siteUrl" -LogLevel "DEBUG"
            
            # Check for subsites
            try {
                $subsites = Invoke-PnPWithRetry -ScriptBlock {
                    Get-PnPSubWeb -Recurse:$false -ErrorAction SilentlyContinue # Get only immediate subsites
                } -Operation "Get-PnPSubWeb for site $siteUrl" -LogName $Log
                $containsSubSites = ($null -ne $subsites -and $subsites.Count -gt 0)
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -ContainsSubSites $containsSubSites
                if ($containsSubSites) { Write-LogEntry -LogName $Log -LogEntryText "Found $($subsites.Count) subsites on site $siteUrl" -LogLevel "INFO" }
            }
            catch { Write-LogEntry -LogName $Log -LogEntryText "Error checking for subsites on site $siteUrl : $_" -LogLevel "ERROR" }

            # Check if it's a Community Site (Yammer integration)
            try {
                $isCommunity = $false
                $navNodes = Invoke-PnPWithRetry -ScriptBlock { Get-PnPNavigationNode } -Operation "Get-PnPNavigationNode for site $siteUrl" -LogName $Log
                # Look for a "Conversations" link pointing to Yammer
                if ($navNodes | Where-Object { $_.Title -eq "Conversations" -and $_.Url -like "*yammer.com*" }) { $isCommunity = $true }
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -IsCommunity $isCommunity
            }
            catch { Write-LogEntry -LogName $Log -LogEntryText "Error checking for Community Site status for $siteUrl : $_" -LogLevel "ERROR" }

            # Get Web properties for site-level user processing (MembersCanShare, AssociatedMemberGroup)
            try {
                $webForSiteLevelUsers = Invoke-PnPWithRetry -ScriptBlock { 
                    Get-PnPWeb -Includes RoleAssignments, AssociatedMemberGroup, MembersCanShare, HasUniqueRoleAssignments 
                } -Operation "Get-PnPWeb for site level users on $siteUrl" -LogName $Log

                $allowMembersEditMembership = $false # Default value
                $membersCanShare = $webForSiteLevelUsers.MembersCanShare # Get MembersCanShare property
                
                # Get AllowMembersEditMembership from the AssociatedMemberGroup if it exists
                if ($webForSiteLevelUsers.AssociatedMemberGroup) {
                    try {
                        Invoke-PnPWithRetry -ScriptBlock { 
                            Get-PnPProperty -ClientObject $webForSiteLevelUsers.AssociatedMemberGroup -Property AllowMembersEditMembership | Out-Null # Load the property
                        } -Operation "Load AllowMembersEditMembership for $siteUrl" -LogName $Log
                        $allowMembersEditMembership = $webForSiteLevelUsers.AssociatedMemberGroup.AllowMembersEditMembership
                    }
                    catch {
                        Write-LogEntry -LogName $Log -LogEntryText "Error getting AssociatedMemberGroup.AllowMembersEditMembership for $siteUrl : $_" -LogLevel "ERROR"
                    }
                }
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -AllowMembersEditMembership $allowMembersEditMembership -MembersCanShare $membersCanShare

                # Process direct site-level role assignments (users/groups with direct permissions on the root web)
                if ($webForSiteLevelUsers.RoleAssignments) {
                    foreach ($roleAssignment in $webForSiteLevelUsers.RoleAssignments) {
                        $member = $null # Initialize member for current role assignment
                        try {
                            # Load the Member property of the role assignment
                            Invoke-PnPWithRetry -ScriptBlock {
                                Get-PnPProperty -ClientObject $roleAssignment -Property Member | Out-Null
                            } -Operation "Load RoleAssignment.Member for site level processing on $siteUrl" -LogName $Log
                            $member = $roleAssignment.Member

                            if ($null -eq $member) { Write-LogEntry -LogName $Log -LogEntryText "Skipping role assignment with null member on $siteUrl" -LogLevel "DEBUG"; continue }

                            $isEveryone = $member.LoginName -like "*spo-grid-all-users*" # Check if member is EEEU

                            # Process if member is a User or EEEU
                            if ($member.PrincipalType -eq [Microsoft.SharePoint.Client.Utilities.PrincipalType]::User -or $isEveryone) {
                                # Skip system accounts and app principals
                                if ($member.LoginName -like "SHAREPOINT\system" -or $member.LoginName -like "*app@sharepoint") { continue }

                                $userNameToStore = $member.Title
                                $userEmailToStore = $member.Email
                                $userLoginToStore = $member.LoginName

                                if ($isEveryone) {
                                    Write-LogEntry -LogName $Log -LogEntryText "Processing EEEU ($userLoginToStore) with direct permissions on $siteUrl." -LogLevel "DEBUG"
                                }
                                # If it's a regular user (has '@'), try to get AAD details for richer info
                                elseif ($member.LoginName -like '*@*') { 
                                    $aadUser = Get-CachedPnPAzureADUser -Identity $member.LoginName -LogName $Log
                                    if ($aadUser) { $userNameToStore = $aadUser.DisplayName; $userEmailToStore = $aadUser.Mail }
                                }

                                # Load RoleDefinitionBindings (permissions) for this member
                                Invoke-PnPWithRetry -ScriptBlock {
                                    Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings | Out-Null
                                } -Operation "Load RoleDefinitionBindings for site level user $($member.LoginName)" -LogName $Log
                                
                                $userRolesCol = [System.Collections.Generic.List[string]]::new()
                                foreach ($roleDef in $roleAssignment.RoleDefinitionBindings) {
                                    if ($roleDef -and $roleDef.Name) {
                                        $userRolesCol.Add($roleDef.Name)
                                    }
                                }

                                # If EEEU, check if permissions are meaningful
                                if ($isEveryone) {
                                    Set-EEEUPresentIfApplicable -SiteUrlToUpdate $siteUrl -Roles $userRolesCol -ContextForLog "Direct Assignment for EEEU"
                                }

                                # FIX: Filter out users/groups with only "Limited Access" for Site Level Users list
                                $hasOnlyLimitedAccessForSiteLevel = $true
                                if ($userRolesCol.Count -eq 0) { $hasOnlyLimitedAccessForSiteLevel = $false } # No roles means not just limited access
                                else {
                                    foreach ($roleNameInCol in $userRolesCol) {
                                        if ($roleNameInCol -ne "Limited Access") {
                                            $hasOnlyLimitedAccessForSiteLevel = $false
                                            break
                                        }
                                    }
                                }
                                
                                # Add to Site Level Users list if they have roles and not *only* Limited Access
                                if ($userRolesCol.Count -gt 0 -and -not $hasOnlyLimitedAccessForSiteLevel) {
                                    $userObject = [PSCustomObject]@{ Name = $userNameToStore; Email = $userEmailToStore; LoginName = $userLoginToStore; Roles = ($userRolesCol | Select-Object -Unique) -join ',' }
                                    $siteCollectionData[$siteUrl]["Site Level Users"].Add($userObject)
                                    Write-LogEntry -LogName $Log -LogEntryText "Added site level principal: $($userObject.Name) with roles: $($userObject.Roles) for $siteUrl" -LogLevel "DEBUG"
                                }
                                elseif ($userRolesCol.Count -gt 0 -and $hasOnlyLimitedAccessForSiteLevel) {
                                    Write-LogEntry -LogName $Log -LogEntryText "Skipping site level principal: $($userNameToStore) as they only have 'Limited Access' for $siteUrl" -LogLevel "DEBUG"
                                }
                            }
                        }
                        catch { Write-LogEntry -LogName $Log -LogEntryText "Error processing a role assignment member ($($member.LoginName)) for site level users on $siteUrl : $_" -LogLevel "ERROR" }
                    }
                }
            }
            catch { Write-LogEntry -LogName $Log -LogEntryText "Error processing site level users/roles for $siteUrl : $_" -LogLevel "ERROR" }
            
            # Get Site Version Policy settings
            try {
                $versionPolicy = Invoke-PnPWithRetry -ScriptBlock { Get-PnPSiteVersionPolicy } -Operation "Get-PnPSiteVersionPolicy for $siteUrl" -LogName $Log
                if ($versionPolicy) {
                    # Handle nulls by setting to -1 (not set)
                    $expireDays = if ($null -eq $versionPolicy.DefaultExpireAfterDays) { -1 } else { [int]$versionPolicy.DefaultExpireAfterDays }
                    $versionLimit = if ($null -eq $versionPolicy.MajorVersionLimit) { -1 } else { [int]$versionPolicy.MajorVersionLimit }
                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -DefaultTrimMode $versionPolicy.DefaultTrimMode -DefaultExpireAfterDays $expireDays -MajorVersionLimit $versionLimit
                }
            }
            catch { Write-LogEntry -LogName $Log -LogEntryText "Error retrieving version policy for $siteUrl : $_" -LogLevel "ERROR" }

            # Get Site Collection Administrators
            try {
                $siteAdmins = Invoke-PnPWithRetry -ScriptBlock { Get-PnPSiteCollectionAdmin } -Operation "Get-PnPSiteCollectionAdmin for $siteUrl" -LogName $Log
                foreach ($admin in $siteAdmins) {
                    if (!$admin -or !$admin.LoginName) { continue } # Skip if admin or login name is null
                    $adminName = $admin.Title; $adminEmail = $admin.Email
                    # If admin is a user (has '@') and is of type User, try to get AAD details
                    if ($admin.LoginName -like '*@*' -and $admin.PrincipalType -eq 'User') {
                        $aadUser = Get-CachedPnPAzureADUser -Identity $admin.LoginName -LogName $Log
                        if ($aadUser) { $adminName = $aadUser.DisplayName; $adminEmail = $aadUser.Mail }
                    }
                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -SiteAdminName $adminName -SiteAdminEmail $adminEmail
                }
            }
            catch { Write-LogEntry -LogName $Log -LogEntryText "Error retrieving site collection admins for $siteUrl : $_" -LogLevel "ERROR" }

            # If the site is Microsoft 365 Group-connected, get Group Owners and Members
            if ($null -ne $siteprops.GroupId -and $siteprops.GroupId -ne [System.Guid]::Empty) {
                try {
                    # Get M365 Group details
                    $aadGroup = Invoke-PnPWithRetry -ScriptBlock { Get-PnPMicrosoft365Group -Identity $siteprops.GroupId } -Operation "Get-PnPMicrosoft365Group for $($siteprops.GroupId)" -LogName $Log
                    if ($aadGroup) { 
                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -AADGroups $aadGroup 

                        $isDynamicGroup = $false
                        # Check if the group is dynamic and store its rule
                        if ($null -ne $aadGroup.MembershipRule) {
                            $isDynamicGroup = $true
                            Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -EntraGroupMembershipRule $aadGroup.MembershipRule
                            Write-LogEntry -LogName $Log -LogEntryText "Group $($siteprops.GroupId) is dynamic. Rule: $($aadGroup.MembershipRule)" -LogLevel "INFO"
                        }

                        # Get M365 Group Owners (always retrieve owners)
                        $groupOwners = Invoke-PnPWithRetry -ScriptBlock { Get-PnPMicrosoft365GroupOwners -Identity $siteprops.GroupId } -Operation "Get M365 Group Owners for $($siteprops.GroupId)" -LogName $Log
                        foreach ($owner in $groupOwners) {
                            # Use UPN if available, otherwise use ID for AAD lookup
                            $ownerIdentity = if (-not [string]::IsNullOrWhiteSpace($owner.UserPrincipalName)) { $owner.UserPrincipalName } else { $owner.Id }
                            $aadOwnerUser = Get-CachedPnPAzureADUser -Identity $ownerIdentity -LogName $Log
                            if ($aadOwnerUser) { Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -EntraGroupOwner $aadOwnerUser.DisplayName -EntraGroupOwnerEmail $aadOwnerUser.Mail }
                        }

                        # Get M365 Group Members conditionally
                        if (-not $isDynamicGroup -or $ExpandDynamicDLs) {
                            if ($isDynamicGroup -and $ExpandDynamicDLs) {
                                Write-LogEntry -LogName $Log -LogEntryText "Expanding dynamic group $($siteprops.GroupId) members as per ExpandDynamicDLs setting." -LogLevel "INFO"
                            }
                            
                            $groupMembers = Invoke-PnPWithRetry -ScriptBlock { Get-PnPMicrosoft365GroupMembers -Identity $siteprops.GroupId } -Operation "Get M365 Group Members for $($siteprops.GroupId)" -LogName $Log
                            foreach ($member in $groupMembers) {
                                $memberIdentity = if (-not [string]::IsNullOrWhiteSpace($member.UserPrincipalName)) { $member.UserPrincipalName } else { $member.Id }
                                $aadMemberUser = Get-CachedPnPAzureADUser -Identity $memberIdentity -LogName $Log
                                if ($aadMemberUser) { Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -EntraGroupMember $aadMemberUser.DisplayName -EntraGroupMemberEmail $aadMemberUser.Mail }
                            }
                        }
                        else {
                            # Is dynamic group AND ExpandDynamicDLs is $false
                            Write-LogEntry -LogName $Log -LogEntryText "Skipping member expansion for dynamic group $($siteprops.GroupId) as per ExpandDynamicDLs setting. Membership rule already stored." -LogLevel "INFO"
                            # Entra Group Members list will remain empty for this site if not expanded; rule is in its own field.
                        }
                    }
                }
                catch { Write-LogEntry -LogName $Log -LogEntryText "Warning: Could not retrieve M365 group info for $($siteprops.GroupId) on $siteUrl : $_" -LogLevel "WARNING" }
            }

            # Get SharePoint Groups and their members
            try {
                $spGroups = Invoke-PnPWithRetry -ScriptBlock { Get-PnPGroup } -Operation "Get-PnPGroup for $siteUrl" -LogName $Log
                ForEach ($spGroup in $spGroups) {
                    if (!$spGroup -or [string]::IsNullOrWhiteSpace($spGroup.Title)) { 
                        Write-LogEntry -LogName $Log -LogEntryText "Skipping SP group with null or empty title on $siteUrl" -LogLevel "WARNING"
                        continue 
                    }
                    $spGroupName = $spGroup.Title; $spGroupRolesString = "" # Initialize roles string
                    
                    # Check for SharingLinks groups to increment "Sharing Links Count"
                    if ($spGroupName -like "SharingLinks*") { 
                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -SPGroupName $spGroupName
                    }

                    # Get roles assigned to this SharePoint group
                    try { 
                        # Determine which web context to use for role assignments (root web or subweb if unique permissions)
                        $currentWebForRoles = if ($webForSiteLevelUsers -and $webForSiteLevelUsers.HasUniqueRoleAssignments) { 
                            $webForSiteLevelUsers # Use already fetched web if it has unique permissions
                        }
                        else { 
                            # Fetch web again if needed, ensuring RoleAssignments are loaded
                            Invoke-PnPWithRetry -ScriptBlock { Get-PnPWeb -Includes RoleAssignments, HasUniqueRoleAssignments } -Operation "Get-PnPWeb for SP Group Roles context $spGroupName" -LogName $Log 
                        }
                        
                        # Find the role assignment for the current SP group
                        $groupRoleAssignment = $currentWebForRoles.RoleAssignments | Where-Object {
                            # Ensure Member property is loaded before accessing LoginName
                            if (-not $_.IsPropertyAvailable("Member")) {
                                Invoke-PnPWithRetry -ScriptBlock { Get-PnPProperty -ClientObject $_ -Property Member | Out-Null } -Operation "Load Member for RA in SP Group $($spGroup.LoginName)" -LogName $Log
                            }
                            $_.Member.LoginName -eq $spGroup.LoginName 
                        }

                        if ($groupRoleAssignment) {
                            # Load RoleDefinitionBindings (permissions) for the group
                            Invoke-PnPWithRetry -ScriptBlock { Get-PnPProperty -ClientObject $groupRoleAssignment -Property RoleDefinitionBindings | Out-Null } -Operation "Load RoleDefBindings for SP Group $spGroupName" -LogName $Log
                            $spGroupRolesString = ($groupRoleAssignment.RoleDefinitionBindings.Name | Select-Object -Unique) -join ',' # Comma-separated list of unique role names
                        }
                    }
                    catch { Write-LogEntry -LogName $Log -LogEntryText "Error retrieving roles for SP group '$spGroupName' on $siteUrl : $_.Exception.Message" -LogLevel "ERROR" } 
                    
                    # Update site data with the SP group name and its roles
                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -SPGroupName $spGroupName -SPGroupRoles $spGroupRolesString

                    # FIX: Determine if the group itself only has "Limited Access"
                    $groupItselfHasOnlyLimitedAccess = $true
                    if ([string]::IsNullOrWhiteSpace($spGroupRolesString)) {
                        # No roles assigned to group directly on the web
                        $groupItselfHasOnlyLimitedAccess = $false # Treat as not "only limited" for member listing purposes.
                        # Or true if no roles means no meaningful access.
                        # Current logic: if group has no roles, it doesn't grant meaningful access via this assignment.
                    }
                    else {
                        $spGroupRolesArrayForCheck = $spGroupRolesString.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                        if ($spGroupRolesArrayForCheck.Count -eq 0) {
                            # Effectively no roles after split/trim
                            $groupItselfHasOnlyLimitedAccess = $false; # Or true, similar to above.
                        }
                        else {
                            foreach ($groupRoleName in $spGroupRolesArrayForCheck) {
                                if ($groupRoleName -ne "Limited Access") {
                                    $groupItselfHasOnlyLimitedAccess = $false
                                    break
                                }
                            }
                        }
                    }

                    # Get members of the current SharePoint Group
                    $spGroupMembers = @()
                    try {
                        if ($spGroup.Id) {
                            # Ensure group ID is available
                            $spGroupMembers = Invoke-PnPWithRetry -ScriptBlock { Get-PnPGroupMember -Identity $spGroup.Id } -Operation "Get members for SP group $spGroupName" -LogName $Log
                        }
                        foreach ($member in $spGroupMembers) {
                            if (!$member -or !$member.LoginName) { continue } # Skip if member or login name is null
                            $spUserLogin = $member.LoginName; $spUserTitle = $member.Title
                            $spUserNameForUpdate = $member.Title; $spUserEmailForUpdate = $member.Email 

                            $isEeeuInSpGroup = $member.LoginName -like "*spo-grid-all-users*" # Check if member is EEEU

                            # If EEEU is a member of this SP group, check if the group's roles are meaningful
                            if ($isEeeuInSpGroup) { 
                                Write-LogEntry -LogName $Log -LogEntryText "EEEU found in SP Group '$spGroupName' (Site: $siteUrl). Roles of group: '$spGroupRolesString'." -LogLevel "DEBUG"
                                $rolesListForEEEUCheckFromGroup = [System.Collections.Generic.List[string]]::new()
                                # Convert comma-separated roles string to list for Set-EEEUPresentIfApplicable
                                $spGroupRolesString.Split(',') | Where-Object { $_.Trim() -ne "" } | ForEach-Object { $rolesListForEEEUCheckFromGroup.Add($_.Trim()) }
                                Set-EEEUPresentIfApplicable -SiteUrlToUpdate $siteUrl -Roles $rolesListForEEEUCheckFromGroup -ContextForLog "member of SP Group '$spGroupName'"
                            }

                            # If member is a regular user, try to get AAD details
                            if ($member.PrincipalType -eq [Microsoft.SharePoint.Client.Utilities.PrincipalType]::User -and $member.LoginName -like '*@*') {
                                $aadUser = Get-CachedPnPAzureADUser -Identity $member.LoginName -LogName $Log
                                if ($aadUser) { $spUserNameForUpdate = $aadUser.DisplayName; $spUserEmailForUpdate = $aadUser.Mail }
                            }
                            
                            # FIX: Add to SP Users list only if the group itself doesn't grant *only* Limited Access
                            if (-not $groupItselfHasOnlyLimitedAccess) {
                                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -AssociatedSPGroup $spGroupName -SPUserName $spUserNameForUpdate -SPUserTitle $spUserTitle -SPUserEmail $spUserEmailForUpdate -SPUserLoginName $spUserLogin
                            }
                            else {
                                Write-LogEntry -LogName $Log -LogEntryText "Skipping user '$($spUserNameForUpdate)' in SP group '$spGroupName' for 'SP Users' list as group only has 'Limited Access' on $siteUrl" -LogLevel "DEBUG"
                            }
                        }
                    }
                    catch { Write-LogEntry -LogName $Log -LogEntryText "Error retrieving members for SP group '$spGroupName' on $siteUrl : $_" -LogLevel "ERROR" }
                }
            }
            catch { Write-LogEntry -LogName $Log -LogEntryText "Error retrieving SP groups for $siteUrl : $_" -LogLevel "ERROR" }

        }
        # Catch errors related to connecting to the specific site or processing site-level details
        catch { Write-LogEntry -LogName $Log -LogEntryText "ERROR: Could not connect to site $siteUrl or process site-level details. $_" -LogLevel "ERROR"; continue } # Continue to next site
    }
    # Catch fatal errors in the main processing block for a site
    catch {
        Write-LogEntry -LogName $Log -LogEntryText "FATAL Error in main processing block for $siteUrl : $_" -LogLevel "ERROR"
        # Ensure basic site info is added if it was fetched before the fatal error
        if ($null -ne $siteUrl -and -not $siteCollectionData.ContainsKey($siteUrl) -and $null -ne $siteprops) {
            Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops 
        }
    }
    
    # Export the collected data for the current site to CSV and remove from memory
    Export-SiteCollectionToCSV -SiteUrl $siteUrl -CsvPath $outputfile -MaxCountForCell $maxcount
    Write-Host "Exported data for site $processedCount/$totalSites to CSV" -ForegroundColor Green

} # End foreach Site

# Disconnect PnP Online session if one exists
if (Get-PnPConnection) {
    Disconnect-PnPOnline
}
Write-LogEntry -LogName $Log -LogEntryText "Disconnected from PnP Online. Script finished."
Write-Host "Script finished. Log file located at: $log" -ForegroundColor Green
Write-Host "Output CSV located at: $outputfile" -ForegroundColor Green
