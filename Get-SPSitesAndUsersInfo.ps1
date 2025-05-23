<#
.SYNOPSIS
    Retrieves detailed information about SharePoint sites, associated SharePoint groups, users, 
    Microsoft 365 groups, and site collection administrators within a tenant, 
    and exports the consolidated data to a CSV file.

.DESCRIPTION
    This script connects to a SharePoint tenant using PnP PowerShell with certificate-based authentication. 
    It retrieves a list of SharePoint sites either from a provided CSV file or directly from the tenant. 
    For each site, it gathers comprehensive details including site properties, SharePoint groups and their roles, SharePoint users, 
    Microsoft 365 group details (if applicable), group owners and members, and site collection administrators. 
    The script consolidates this information into a structured format and exports it to a CSV file for reporting and auditing purposes.

.PARAMETER tenantname
    Specifies the SharePoint tenant name (e.g., "contoso" for contoso-admin.sharepoint.com).

.PARAMETER appID
    Specifies the Azure AD (Entra) Application ID used for authentication.

.PARAMETER thumbprint
    Specifies the certificate thumbprint used for authentication.

.PARAMETER tenant
    Specifies the Azure AD (Entra) Tenant ID.

.PARAMETER inputfile
    (Optional) Path to a CSV file containing a list of SharePoint site URLs to process. If not provided or not found, the script retrieves all sites from the tenant.

.PARAMETER Debug
    (Optional) Set to $True to include debug-level messages in the log file. Default is $True.

.OUTPUTS
    CSV file containing detailed information about each processed SharePoint site, including:
        - Site URL and properties (Owner, Template, Sharing settings, Information Barrier settings, Teams connection status, etc.)
        - SharePoint groups and their assigned roles
        - SharePoint users and their associated groups
        - Microsoft 365 group details (Display Name, Alias, Access Type, Creation Date)
        - Microsoft 365 group owners and members
        - Site collection administrators
        - Indicators for sharing links and "EEEU Present" status

.NOTES

    Authors: Mike Lee
    Date: 5/23/25
    Script includes throttling handling for SharePoint Online

    Requirements:
        - PnP.PowerShell module installed (Tested with PNP 2.12.0)
        - PowerShell 7.4 or higer
        - Appropriate permissions granted to the Azure AD application
            - Microsoft Graph| Application | Directory.Read.All
            - SharePoint |Application | Sites.FullControl.All
        - Certificate-based authentication configured
        

Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

.EXAMPLE
    .\Get-SPSitesAndUsersInfo.ps1

    Runs the script using default parameters, retrieves all SharePoint sites from the tenant, and exports the detailed information to a CSV file.

.EXAMPLE
    .\Get-SPSitesAndUsersInfo.ps1 -inputfile "C:\temp\sitelist.csv"

    Runs the script using a provided CSV file containing specific SharePoint site URLs to process.

.EXAMPLE
    .\Get-SPSitesAndUsersInfo.ps1 -Debug $False

    Runs the script with debug logging disabled, only writing information log entries to the log file.

#>
# Set Variables
$tenantname = "m365x61250205" #This is your tenant name
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"  #This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9" #This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51" #This is your Tenant ID
$Debug = $false # Set to $True to include debug-level messages in the log file, $False to exclude them

#Initialize Parameters - Do not change
$sites = @()
$inputfile = $null
$outputfile = $null
$log = $null
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$maxRetries = 5  # Maximum number of retry attempts
$initialRetryDelay = 5  # Initial retry delay in seconds

#Input / Output and Log Files
$inputfile = "C:\temp\sitelist-m365x61250205.csv" #This is the input file with list of sites to process. If not provided, all sites will be processed.
$outputfile = "$env:TEMP\" + 'SPSites_and_Users_Info_' + $date + '_' + "output.csv"
$log = "$env:TEMP\" + 'SPSites_and_Users_Info_' + $date + '_' + "logfile.log"

#This is the logging function
Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $LogLevel = "INFO"  # Default log level is INFO
    )
    if ($LogName -ne $null) {
        # Skip DEBUG level messages if Debug is set to False
        if ($LogLevel -eq "DEBUG" -and $Debug -eq $False) {
            return
        }
        
        # log the date and time in the text file along with the data passed
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : [$LogLevel] $LogEntryText" | Out-File -FilePath $LogName -append;
    }
}

# Function to handle throttling with exponential backoff
Function Invoke-PnPWithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock] $ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [string] $Operation = "PnP Operation",
        
        [Parameter(Mandatory = $false)]
        [int] $MaxRetries = 5,
        
        [Parameter(Mandatory = $false)]
        [int] $InitialRetryDelay = 5,
        
        [Parameter(Mandatory = $false)]
        [string] $LogName
    )
    
    $retryCount = 0
    $success = $false
    $result = $null
    $retryDelay = $InitialRetryDelay
    
    do {
        try {
            $result = & $ScriptBlock
            $success = $true
            return $result
        }
        catch {
            $exceptionDetails = $_.Exception.ToString()
            
            # Check for throttling-related exceptions
            if (($exceptionDetails -like "*429*") -or 
                ($exceptionDetails -like "*throttl*") -or 
                ($exceptionDetails -like "*too many requests*") -or
                ($exceptionDetails -like "*request limit exceeded*")) {
                
                $retryCount++
                
                # Check if we've hit max retries
                if ($retryCount -ge $MaxRetries) {
                    Write-LogEntry -LogName $Log -LogEntryText "Max retries ($MaxRetries) reached for $Operation. Giving up." -LogLevel "ERROR"
                    throw $_
                }
                
                # Parse Retry-After header if available
                $retryAfterValue = $null
                if ($_.Exception.Response -and $_.Exception.Response.Headers -and $_.Exception.Response.Headers["Retry-After"]) {
                    $retryAfterValue = [int]$_.Exception.Response.Headers["Retry-After"]
                    $retryDelay = $retryAfterValue
                    Write-LogEntry -LogName $Log -LogEntryText "Throttling detected for $Operation. Server requested retry after $retryAfterValue seconds." -LogLevel "WARNING"
                }
                else {
                    # Use exponential backoff if no Retry-After header
                    $retryDelay = [Math]::Min(30, $retryDelay * 2)
                    Write-LogEntry -LogName $Log -LogEntryText "Throttling detected for $Operation. Using exponential backoff: waiting $retryDelay seconds before retry $retryCount of $MaxRetries." -LogLevel "WARNING"
                }
                
                Write-Host "Throttling detected for $Operation. Waiting $retryDelay seconds before retry $retryCount of $MaxRetries." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryDelay
            }
            else {
                # Not a throttling error, just throw it
                throw $_
            }
        }
    } while (-not $success -and $retryCount -lt $MaxRetries)
}

# Define the connection parameters for reuse
$connectionParams = @{
    ClientId      = $appID
    Thumbprint    = $thumbprint
    Tenant        = $tenant
    WarningAction = 'SilentlyContinue'
}

#Connect to Admin Center initially
try {
    $adminUrl = 'https://' + $tenantname + '-admin.sharepoint.com'
    
    # Use the retry function for the connection
    Invoke-PnPWithRetry -ScriptBlock { 
        Connect-PnPOnline -Url $adminUrl @connectionParams 
    } -Operation "Connect to SharePoint Admin Center" -LogName $Log
    
    Write-LogEntry -LogName $Log -LogEntryText "Successfully connected to SharePoint Admin Center: $adminUrl"
}
catch {
    Write-Host "Error connecting to SharePoint Admin Center ($adminUrl): $_" -ForegroundColor Red
    Write-LogEntry -LogName $Log -LogEntryText "Error connecting to SharePoint Admin Center ($adminUrl): $_" -LogLevel "ERROR"
    exit
}

# Get Site List
if ($inputfile -and (Test-Path -Path $inputfile)) {
    try {
        $sites = Import-csv -path $inputfile -Header 'URL'
        Write-LogEntry -LogName $Log -LogEntryText "Using sites from input file: $inputfile"
        Write-Host "Reading sites from input file: $inputfile" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Error reading input file '$inputfile': $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error reading input file '$inputfile': $_" -LogLevel "ERROR"
        exit
    }
}
else {
    Write-Host "Getting site list from tenant (this might take a while)..." -ForegroundColor Yellow
    Write-LogEntry -LogName $Log -LogEntryText "Getting sites using Get-PnPTenantSite (no input file specified or found)"
    try {
        # Ensure we are connected to Admin Center before this call
        Invoke-PnPWithRetry -ScriptBlock { 
            Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop 
        } -Operation "Connect to SharePoint Admin Center (before Get-PnPTenantSite)" -LogName $Log
        
        # Use retry function for getting tenant sites which is prone to throttling
        $sites = Invoke-PnPWithRetry -ScriptBlock { 
            # Excludes OneDrive by default and Redirect Sites
            Get-PnPTenantSite  -Filter { 'Url' -notlike '-my.sharepoint.com' } | Where-Object { $_.Template -ne 'RedirectSite#0' }
        } -Operation "Get-PnPTenantSite" -LogName $Log
        
        Write-Host "Found $($sites.Count) sites." -ForegroundColor Green
        Write-LogEntry -LogName $Log -LogEntryText "Retrieved $($sites.Count) sites using Get-PnPTenantSite."
    }
    catch {
        Write-Host "Error getting site list from tenant: $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error getting site list from tenant: $_" -LogLevel "ERROR"
        exit
    }
}

# Initialize a hashtable to store site collection data (keyed by URL)
$siteCollectionData = @{}

# Modified function to handle consolidated site data
Function Update-SiteCollectionData {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [object] $SiteProperties,
        [string] $SPGroupName = "",
        [string] $SPGroupRoles = "",
        # --- Parameters for SP User ---
        [string] $AssociatedSPGroup = "", # Track which SP group the user is from
        [string] $SPUserName = "",
        [string] $SPUserTitle = "",
        [string] $SPUserEmail = "",
        [string] $SPUserLoginName = "", # Added to check for *spo-grid-all-users*
        # --- Parameters for Entra Group ---
        [string] $EntraGroupOwner = "",
        [string] $EntraGroupOwnerEmail = "",
        [string] $EntraGroupMember = "",
        [string] $EntraGroupMemberEmail = "",
        [object] $AADGroups = $null,
        # --- Parameters for Site Collection Admins ---
        [string] $SiteAdminName = "",
        [string] $SiteAdminEmail = "",
        # --- Parameters for Version Policy ---
        [string] $DefaultTrimMode = "",
        [int] $DefaultExpireAfterDays = -1,  # Changed from 0 to -1 as default
        [int] $MajorVersionLimit = -1,        # Changed from 0 to -1 as default
        # --- Parameter for Community Site ---
        [bool] $IsCommunity = $false,        # Parameter to indicate if site is a Community Site
        # --- Parameters for Member Group Settings ---
        [bool] $AllowMembersEditMembership = $false,
        [bool] $MembersCanShare = $false,
        # --- Parameter for Subsites Detection ---
        [bool] $ContainsSubSites = $false    # Parameter to indicate if site contains subsites
    )

    # Create site entry if it doesn't exist
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
            # Version Policy Settings - Set default values initially
            "DefaultTrimMode"                = $DefaultTrimMode
            "DefaultExpireAfterDays"         = $DefaultExpireAfterDays
            "MajorVersionLimit"              = $MajorVersionLimit
            # Community Site Status
            "Community Site"                 = $IsCommunity
            # Member Group Settings
            "AllowMembersEditMembership"     = $AllowMembersEditMembership
            "MembersCanShare"                = $MembersCanShare
            # Subsites Detection
            "Contains SubSites"              = $ContainsSubSites
            # Site-specific lists
            "SP Groups On Site"              = [System.Collections.Generic.List[string]]::new()
            "SP Group Roles Per Group"       = [System.Collections.Generic.Dictionary[string, string]]::new()
            "SP Users"                       = [System.Collections.Generic.List[PSObject]]::new() # Stores {AssociatedSPGroup, Name, Title, Email}
            "Entra Group Owners"             = [System.Collections.Generic.List[PSObject]]::new() # Stores {Name, Email}
            "Entra Group Members"            = [System.Collections.Generic.List[PSObject]]::new() # Stores {Name, Email}
            "Entra Group Details"            = $null
            "Site Collection Admins"         = [System.Collections.Generic.List[PSObject]]::new() # Stores {Name, Email}
            "Site Level Users"               = [System.Collections.Generic.List[PSObject]]::new() # Stores {Name, Email, LoginName, Roles}
            "Has Sharing Links"              = $false #Property to track if sharing links are being used
            "EEEU Present"                   = $false #"Shared With Everyone" =  "EEEU Present"
        }
    }
    else {
        # Update Community Site status if provided
        if ($PSBoundParameters.ContainsKey('IsCommunity')) {
            $siteCollectionData[$SiteUrl]["Community Site"] = $IsCommunity
        }
        
        # Update Subsites status if provided
        if ($PSBoundParameters.ContainsKey('ContainsSubSites')) {
            $siteCollectionData[$SiteUrl]["Contains SubSites"] = $ContainsSubSites
        }
        
        # If the site entry exists and version policy parameters are provided, update them
        if (-not [string]::IsNullOrEmpty($DefaultTrimMode)) {
            $siteCollectionData[$SiteUrl]["DefaultTrimMode"] = $DefaultTrimMode
        }
        
        # Updated to handle zero values correctly
        if ($DefaultExpireAfterDays -ge 0) {
            $siteCollectionData[$SiteUrl]["DefaultExpireAfterDays"] = $DefaultExpireAfterDays
        }
        
        # Updated to handle zero values correctly
        if ($MajorVersionLimit -ge 0) {
            $siteCollectionData[$SiteUrl]["MajorVersionLimit"] = $MajorVersionLimit
        }

        # Update Member Group Settings if provided
        if ($PSBoundParameters.ContainsKey('AllowMembersEditMembership')) {
            $siteCollectionData[$SiteUrl]["AllowMembersEditMembership"] = $AllowMembersEditMembership
        }
        
        if ($PSBoundParameters.ContainsKey('MembersCanShare')) {
            $siteCollectionData[$SiteUrl]["MembersCanShare"] = $MembersCanShare
        }
    }

    # Check for SharingLinks groups
    if (-not [string]::IsNullOrWhiteSpace($SPGroupName) -and $SPGroupName -like "SharingLinks*") {
        $siteCollectionData[$SiteUrl]["Has Sharing Links"] = $true
    }

    # Check for "shared with everyone" through SP users - ONLY WHEN EXPLICITLY FINDING the user
    if (-not [string]::IsNullOrWhiteSpace($SPUserLoginName) -and $SPUserLoginName -like "*spo-grid-all-users*") {
        Write-LogEntry -LogName $Log -LogEntryText "Found LoginName with spo-grid-all-users: $SPUserLoginName - Setting 'EEEU Present' to TRUE for site $SiteUrl" -LogLevel "DEBUG"
        $siteCollectionData[$SiteUrl]["EEEU Present"] = $true
    }

    # Add Site Collection Admin information (checking for duplicates based on email)
    if (-not [string]::IsNullOrWhiteSpace($SiteAdminEmail)) {
        $checkEmail = $SiteAdminEmail
        $existingAdmin = $siteCollectionData[$SiteUrl]["Site Collection Admins"].Find({ param($a) $a.Email -eq $checkEmail })
        if ($null -eq $existingAdmin) {
            $adminObject = [PSCustomObject]@{ Name = $SiteAdminName; Email = $SiteAdminEmail }
            $siteCollectionData[$SiteUrl]["Site Collection Admins"].Add($adminObject)
        }
    }

    # Add AAD Group information if available (only once per site) - This will now be handled by the caller
    if ($AADGroups) {
        $siteCollectionData[$SiteUrl]["Entra Group Details"] = [PSCustomObject]@{
            DisplayName = $AADGroups.DisplayName; Alias = $AADGroups.MailNickname
            AccessType = $AADGroups.Visibility; WhenCreated = $AADGroups.CreatedDateTime
        }
    }

    # Add SP Group and its Roles if provided and not already present for this site
    if (-not [string]::IsNullOrWhiteSpace($SPGroupName)) {
        if (-not $siteCollectionData[$SiteUrl]["SP Groups On Site"].Contains($SPGroupName)) {
            $siteCollectionData[$SiteUrl]["SP Groups On Site"].Add($SPGroupName)
        }
        if (-not [string]::IsNullOrWhiteSpace($SPGroupRoles)) {
            if ($siteCollectionData[$SiteUrl]["SP Group Roles Per Group"].ContainsKey($SPGroupName)) {
                $siteCollectionData[$SiteUrl]["SP Group Roles Per Group"][$SPGroupName] = $SPGroupRoles
            }
            else {
                $siteCollectionData[$SiteUrl]["SP Group Roles Per Group"].Add($SPGroupName, $SPGroupRoles)
            }
        }
    }

    # Add SharePoint User information (associated with a specific SP group)
    if (-not [string]::IsNullOrWhiteSpace($SPUserName)) {
        $userObject = [PSCustomObject]@{
            AssociatedSPGroup = $AssociatedSPGroup # Store the group name
            Name              = $SPUserName
            Title             = $SPUserTitle
            Email             = $SPUserEmail
        }
        $siteCollectionData[$SiteUrl]["SP Users"].Add($userObject)
    }

    # Add Entra Group Owner information (checking for duplicates based on email)
    if (-not [string]::IsNullOrWhiteSpace($EntraGroupOwnerEmail)) {
        $checkEmail = $EntraGroupOwnerEmail
        $existingOwner = $siteCollectionData[$SiteUrl]["Entra Group Owners"].Find({ param($o) $o.Email -eq $checkEmail })
        if ($null -eq $existingOwner) {
            $ownerObject = [PSCustomObject]@{ Name = $EntraGroupOwner; Email = $EntraGroupOwnerEmail }
            $siteCollectionData[$SiteUrl]["Entra Group Owners"].Add($ownerObject)
        }
    }

    # Add Entra Group Member information (checking for duplicates based on email)
    if (-not [string]::IsNullOrWhiteSpace($EntraGroupMemberEmail)) {
        $checkEmail = $EntraGroupMemberEmail
        $existingMember = $siteCollectionData[$SiteUrl]["Entra Group Members"].Find({ param($m) $m.Email -eq $checkEmail })
        if ($null -eq $existingMember) {
            $memberObject = [PSCustomObject]@{ Name = $EntraGroupMember; Email = $EntraGroupMemberEmail }
            $siteCollectionData[$SiteUrl]["Entra Group Members"].Add($memberObject)
        }
    }
}

# --- Main Processing Loop ---
$totalSites = $sites.Count
$processedCount = 0

# Create CSV with headers first
$csvHeaders = "URL,Owner,IB Mode,IB Segment,Group ID,RelatedGroupId,IsHubSite,Template,SiteDefinedSharingCapability," + 
"SharingCapability,DisableCompanyWideSharingLinks,AllowMembersEditMembership,MembersCanShare,Custom Script Allowed,IsTeamsConnected,IsTeamsChannelConnected," + 
"TeamsChannelType,StorageQuota (MB),StorageUsageCurrent (MB),LockState,LastContentModifiedDate,ArchiveState," + 
"DefaultTrimMode,DefaultExpireAfterDays,MajorVersionLimit,Entra Group Alias," + 
"Entra Group AccessType,Entra Group WhenCreated,Has Sharing Links," + 
"EEEU Present,Community Site,Contains SubSites,SP Groups On Site,SP Groups Roles,Site Collection Admins (Name <Email>)," +
"Site Level Users (Name <Email> [Roles]), SP Users (Group: Name <Email>),Entra Group Owners (Name <Email>),Entra Group Members (Name <Email>)"

# Create the CSV file with headers
Set-Content -Path $outputfile -Value $csvHeaders -Encoding UTF8
Write-Host "Created output file with headers: $outputfile" -ForegroundColor Green
Write-LogEntry -LogName $Log -LogEntryText "Created output file with headers: $outputfile"

# Function to export a single site collection record to CSV
function Export-SiteCollectionToCSV {
    param(
        [string] $SiteUrl,
        [string] $CsvPath
    )
    
    $siteData = $siteCollectionData[$SiteUrl]
    if (-not $siteData) {
        Write-LogEntry -LogName $Log -LogEntryText "Error: No data found for site $SiteUrl when attempting to export" -LogLevel "ERROR"
        return
    }

    # --- Format the combined strings ---
    # SP Users: "GroupName:Name <Email>"
    $spUsersFormatted = ($siteData."SP Users" | ForEach-Object {
            $emailStr = $_.Email | Out-String -NoNewline # Handle potential $null email
            "$($_.AssociatedSPGroup):$($_.Name) <$emailStr>"
        }) -join ';'
  
    # Entra Owners: "Name <Email>"
    $entraOwnersFormatted = ($siteData."Entra Group Owners" | ForEach-Object {
            $emailStr = $_.Email | Out-String -NoNewline
            "$($_.Name) <$emailStr>"
        }) -join ';'

    # Entra Members: "Name <Email>"
    $entraMembersFormatted = ($siteData."Entra Group Members" | ForEach-Object {
            $emailStr = $_.Email | Out-String -NoNewline
            "$($_.Name) <$emailStr>"
        }) -join ';'
    
    # Site Collection Admins: "Name <Email>"
    $siteAdminsFormatted = ($siteData."Site Collection Admins" | ForEach-Object {
            $emailStr = $_.Email | Out-String -NoNewline
            "$($_.Name) <$emailStr>"
        }) -join ';'
    
    # Site Level Users: "Name <Email> [Roles]"
    $siteLevelUsersFormatted = ($siteData."Site Level Users" | ForEach-Object {
            $emailStr = $_.Email | Out-String -NoNewline
            "$($_.Name) <$emailStr> [$($_.Roles)]"
        }) -join ';'

    # --- Create the export object with combined columns ---
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
        "Has Sharing Links"                       = if ($siteData."Has Sharing Links") { "True" } else { "False" }
        "EEEU Present"                            = if ($siteData."EEEU Present") { "True" } else { "False" }
        "Community Site"                          = if ($siteData."Community Site") { "True" } else { "False" }
        "Contains SubSites"                       = if ($siteData."Contains SubSites") { "True" } else { "False" }
        "AllowMembersEditMembership"              = if ($siteData."AllowMembersEditMembership") { "True" } else { "False" }
        "MembersCanShare"                         = if ($siteData."MembersCanShare") { "True" } else { "False" }
        "SP Groups On Site"                       = ($siteData."SP Groups On Site" -join ';')
        "SP Groups Roles"                         = ($siteData."SP Group Roles Per Group".Values | Select-Object -Unique | Where-Object { $_ }) -join ';'
        "Site Collection Admins (Name <Email>)"   = $siteAdminsFormatted
        "SP Users (Group: Name <Email>)"          = $spUsersFormatted       # Combined SP User Info
        "Site Level Users (Name <Email> [Roles])" = $siteLevelUsersFormatted # Combined Site Level Users with Roles
        "Entra Group Owners (Name <Email>)"       = $entraOwnersFormatted   # Combined Owner Info
        "Entra Group Members (Name <Email>)"      = $entraMembersFormatted  # Combined Member Info
    }

    # Export this item as a single line to CSV (append mode)
    try {
        $exportItem | Export-Csv -Path $CsvPath -NoTypeInformation -Append -Encoding UTF8
        Write-LogEntry -LogName $Log -LogEntryText "Successfully wrote data for site $SiteUrl to CSV" -LogLevel "DEBUG"
        
        # Remove the site data from the hashtable to free memory
        $siteCollectionData.Remove($SiteUrl)
    }
    catch {
        Write-Host "Error writing site data ($SiteUrl) to CSV '$CsvPath': $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error writing site data ($SiteUrl) to CSV '$CsvPath': $_" -LogLevel "ERROR"
    }
}

foreach ($site in $sites) {
    $processedCount++
    $siteUrl = $site.Url
    Write-Host "Processing site $processedCount/$totalSites : $siteUrl" -ForegroundColor Cyan
    Write-LogEntry -LogName $Log -LogEntryText "Processing site $processedCount/$totalSites : $siteUrl"

    $siteprops = $null
    $AADGroups = $null # M365 Group details
    $groupmembersRaw = $null # M365 Group Members
    $groupownersRaw = $null # M365 Group Owners
    $currentPnPConnection = $null # To hold the site-specific connection if successful
    $containsSubSites = $false # Default value for SubSites detection

    try {
        # Get Site Properties using the Admin connection context
        Invoke-PnPWithRetry -ScriptBlock { 
            Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop 
        } -Operation "Connect to Admin URL for site $siteUrl" -LogName $Log
        
        $siteprops = Invoke-PnPWithRetry -ScriptBlock { 
            Get-PnPTenantSite -Identity $siteUrl | Select-Object Url, Owner, InformationBarrierMode, InformationBarrierSegments, GroupId, RelatedGroupId, IsHubSite, Template, SiteDefinedSharingCapability, SharingCapability, DisableCompanyWideSharingLinks, DenyAddAndCustomizePages, IsTeamsConnected, IsTeamsChannelConnected, TeamsChannelType, StorageQuota, StorageUsageCurrent, LockState, LastContentModifiedDate, ArchiveState
        } -Operation "Get-PnPTenantSite for $siteUrl" -LogName $Log

        if ($null -eq $siteprops) { Write-LogEntry -LogName $Log -LogEntryText "Failed to retrieve properties for site $siteUrl. Skipping." -LogLevel "ERROR"; continue }

        # Initialize site data with basic properties
        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops

        # --- Connect to the specific site ---
        try {
            Write-LogEntry -LogName $Log -LogEntryText "Connecting to specific site: $siteUrl" -LogLevel "DEBUG"
            
            $currentPnPConnection = Invoke-PnPWithRetry -ScriptBlock { 
                Connect-PnPOnline -Url $siteUrl @connectionParams -ErrorAction Stop 
            } -Operation "Connect to site $siteUrl" -LogName $Log
            
            Write-LogEntry -LogName $Log -LogEntryText "Successfully connected to specific site: $siteUrl" -LogLevel "DEBUG"
            
            # Check for subsites
            try {
                Write-LogEntry -LogName $Log -LogEntryText "Checking for subsites on site $siteUrl" -LogLevel "DEBUG"
                
                $subsites = Invoke-PnPWithRetry -ScriptBlock {
                    Get-PnPSubWeb -Recurse:$false -ErrorAction SilentlyContinue
                } -Operation "Get-PnPSubWeb for site $siteUrl" -LogName $Log
                
                if ($null -ne $subsites -and $subsites.Count -gt 0) {
                    $containsSubSites = $true
                    Write-LogEntry -LogName $Log -LogEntryText "Found $($subsites.Count) subsites on site $siteUrl" -LogLevel "DEBUG"
                }
                else {
                    Write-LogEntry -LogName $Log -LogEntryText "No subsites found on site $siteUrl" -LogLevel "DEBUG"
                }
                
                # Update site data with SubSites status
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -ContainsSubSites $containsSubSites
            }
            catch {
                Write-LogEntry -LogName $Log -LogEntryText "Error checking for subsites on site $siteUrl : $_" -LogLevel "ERROR"
                # Ensure SubSites is set to false if an error occurs
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -ContainsSubSites $false
            }
            
            # Check if this is a Community Site by examining navigation nodes for Yammer link
            try {
                Write-LogEntry -LogName $Log -LogEntryText "Checking if site $siteUrl is a Community Site" -LogLevel "DEBUG"
                
                $isCommunity = $false
                $navNodes = Invoke-PnPWithRetry -ScriptBlock {
                    Get-PnPNavigationNode
                } -Operation "Get-PnPNavigationNode for site $siteUrl" -LogName $Log
                
                # Look for "Conversations" node with a URL containing "yammer.com"
                $yammerNode = $navNodes | Where-Object { 
                    $_.Title -eq "Conversations" -and $_.Url -like "*yammer.com*" 
                }
                
                if ($null -ne $yammerNode) {
                    Write-LogEntry -LogName $Log -LogEntryText "Found Yammer integration on site $($siteUrl): $($yammerNode.Url)" -LogLevel "DEBUG"
                    $isCommunity = $true
                }
                
                # Update site data with Community Site status
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -IsCommunity $isCommunity
                Write-LogEntry -LogName $Log -LogEntryText "Set Community Site status to $isCommunity for site $siteUrl" -LogLevel "DEBUG"
            }
            catch {
                Write-LogEntry -LogName $Log -LogEntryText "Error checking for Community Site status for $siteUrl : $_" -LogLevel "ERROR"
                # Ensure Community Site is set to false if an error occurs
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -IsCommunity $false
            }
            
            # Check for "spo-grid-all-users" at the site collection level - with additional debugging
            try {
                Write-LogEntry -LogName $Log -LogEntryText "Checking for 'Everyone' user at site collection level for $siteUrl" -LogLevel "DEBUG"
                
                $allSiteUsers = Invoke-PnPWithRetry -ScriptBlock {
                    Get-PnPUser -WithRightsAssigned
                } -Operation "Get-PnPUser for site collection $siteUrl" -LogName $Log
                
                # Add debug logging to see what users are being returned
                Write-LogEntry -LogName $Log -LogEntryText "Found $($allSiteUsers.Count) users at site collection level for $siteUrl" -LogLevel "DEBUG"
                
                # Check if any user has "spo-grid-all-users" in their login name - with specific filter and debug
                $everyoneUser = $allSiteUsers | Where-Object { 
                    $hasPattern = $_.LoginName -like "*spo-grid-all-users*"
                    if ($hasPattern -eq 'True') {
                        Write-LogEntry -LogName $Log -LogEntryText "FOUND MATCH: User $($_.Title) with login $($_.LoginName) matches spo-grid-all-users pattern" -LogLevel "DEBUG"
                    }
                    return $hasPattern
                }
                
                if ($null -ne $everyoneUser -and $everyoneUser.Count -gt 0) {
                    Write-LogEntry -LogName $Log -LogEntryText "Found 'Everyone' user (spo-grid-all-users) at site collection level on $siteUrl" -LogLevel "DEBUG"
                    
                    # First initialize site data if not already done
                    if (-not $siteCollectionData.ContainsKey($siteUrl)) {
                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops
                    }
                    
                    # Directly update the hashtable for this specific site
                    $siteCollectionData[$siteUrl]["EEEU Present"] = $true
                    Write-LogEntry -LogName $Log -LogEntryText "EXPLICITLY Setting 'EEEU Present' to TRUE for site $siteUrl" -LogLevel "DEBUG"
                }
                else {
                    Write-LogEntry -LogName $Log -LogEntryText "No 'Everyone' user found at site collection level for $siteUrl" -LogLevel "DEBUG"
                    
                    # Ensure this site isn't incorrectly flagged
                    if ($siteCollectionData.ContainsKey($siteUrl)) {
                        $siteCollectionData[$siteUrl]["EEEU Present"] = $false
                    }
                }
                
                # Process site level users with assigned permissions
                if ($allSiteUsers -and $allSiteUsers.Count -gt 0) {
                    Write-LogEntry -LogName $Log -LogEntryText "Processing $($allSiteUsers.Count) site level users with direct permissions on $siteUrl" -LogLevel "DEBUG"
                    
                    # Initialize the site level users collection if needed
                    if (-not $siteCollectionData.ContainsKey($siteUrl)) {
                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops
                    }
                    if ($null -eq $siteCollectionData[$siteUrl]["Site Level Users"]) {
                        $siteCollectionData[$siteUrl]["Site Level Users"] = [System.Collections.Generic.List[PSObject]]::new()
                    }
                    
                    # Get Web object once for permission checks
                    $web = Invoke-PnPWithRetry -ScriptBlock { 
                        Get-PnPWeb -Includes RoleAssignments, AssociatedMemberGroup, MembersCanShare 
                    } -Operation "Get-PnPWeb with RoleAssignments for site level users on $siteUrl" -LogName $Log
                    
                    # Capture AllowMembersEditMembership and MembersCanShare properties
                    $allowMembersEditMembership = $false
                    $membersCanShare = $false
                    
                    if ($null -ne $web.MembersCanShare) {
                        $membersCanShare = $web.MembersCanShare
                        Write-LogEntry -LogName $Log -LogEntryText "MembersCanShare value for site $($siteUrl): $membersCanShare" -LogLevel "DEBUG"
                    }
                    
                    if ($null -ne $web.AssociatedMemberGroup) {
                        # Need to load AssociatedMemberGroup.AllowMembersEditMembership explicitly
                        try {
                            $memberGroup = Invoke-PnPWithRetry -ScriptBlock { 
                                Get-PnPProperty -ClientObject $web -Property AssociatedMemberGroup 
                            } -Operation "Get-PnPProperty AssociatedMemberGroup for $siteUrl" -LogName $Log
                            
                            if ($null -ne $memberGroup) {
                                # Load AllowMembersEditMembership property
                                Invoke-PnPWithRetry -ScriptBlock { 
                                    Get-PnPProperty -ClientObject $memberGroup -Property AllowMembersEditMembership | Out-Null
                                } -Operation "Get-PnPProperty AllowMembersEditMembership for $siteUrl" -LogName $Log
                                
                                $allowMembersEditMembership = $memberGroup.AllowMembersEditMembership
                                Write-LogEntry -LogName $Log -LogEntryText "AllowMembersEditMembership value for site $($siteUrl): $allowMembersEditMembership" -LogLevel "DEBUG"
                            }
                        }
                        catch {
                            Write-LogEntry -LogName $Log -LogEntryText "Error getting AssociatedMemberGroup.AllowMembersEditMembership for $($siteUrl): $_" -LogLevel "ERROR"
                        }
                    }
                    
                    # Update site data with the AllowMembersEditMembership properties
                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -AllowMembersEditMembership $allowMembersEditMembership -MembersCanShare $membersCanShare
                    
                    foreach ($siteUser in $allSiteUsers) {
                        try {
                            # Special handling for "Everyone" security group
                            $isEveryone = $siteUser.LoginName -like "*spo-grid-all-users*"
                            
                            # Process both individual users and the Everyone security group
                            if ($siteUser.PrincipalType -ne 'User' -and -not $isEveryone) {
                                Write-LogEntry -LogName $Log -LogEntryText "Skipping non-User principal (except Everyone): '$($siteUser.Title)' ($($siteUser.PrincipalType)) on $siteUrl" -LogLevel "DEBUG"
                                continue
                            }
                            
                            $userName = $siteUser.Title
                            $userEmail = $siteUser.Email
                            $userLogin = $siteUser.LoginName
                            
                            # Skip system accounts or special accounts
                            if ($userLogin -like "SHAREPOINT\system" -or 
                                $userLogin -like "*app@sharepoint") {
                                Write-LogEntry -LogName $Log -LogEntryText "Skipping system/special account: $userLogin" -LogLevel "DEBUG"
                                continue
                            }
                            
                            # Enhanced logging for Everyone group
                            if ($isEveryone) {
                                Write-LogEntry -LogName $Log -LogEntryText "Processing 'Everyone' security group ($userLogin) - will be included in site level users" -LogLevel "DEBUG"
                            }
                            
                            # Get additional info from Azure AD if it's a user account
                            if ($userLogin -like '*@*') {
                                try {
                                    $aadUser = Invoke-PnPWithRetry -ScriptBlock { 
                                        Get-PnPAzureADUser -Identity $userLogin -ErrorAction SilentlyContinue 
                                    } -Operation "Get-PnPAzureADUser for site user $userLogin" -LogName $Log
                                    
                                    if ($aadUser) { 
                                        $userName = $aadUser.DisplayName
                                        $userEmail = $aadUser.Mail
                                    }
                                }
                                catch { 
                                    Write-LogEntry -LogName $Log -LogEntryText "Warning: Getting AAD User info for site user '$userLogin' failed: $_" -LogLevel "WARNING" 
                                }
                            }
                            
                            # Get user's roles
                            $userRoles = @()
                            $hasDirectPermissions = $false
                            
                            foreach ($roleAssignment in $web.RoleAssignments) {
                                try {
                                    $member = Invoke-PnPWithRetry -ScriptBlock { 
                                        Get-PnPProperty -ClientObject $roleAssignment -Property Member 
                                    } -Operation "Get RoleAssignment Member for site user/group $userLogin" -LogName $Log
                                    
                                    # Check if this role assignment is for our current user/group
                                    if ($member -and $member.LoginName -eq $userLogin) {
                                        $hasDirectPermissions = $true
                                        
                                        $roleDefinitions = Invoke-PnPWithRetry -ScriptBlock { 
                                            Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings 
                                        } -Operation "Get RoleDefinitionBindings for site user/group $userLogin" -LogName $Log
                                        
                                        foreach ($roleDef in $roleDefinitions) {
                                            if ($roleDef -and $roleDef.Name) {
                                                # Include all permission types for Everyone, but skip limited access for regular users
                                                if ($isEveryone -or $roleDef.Name -ne "Limited Access") {
                                                    $userRoles += $roleDef.Name
                                                }
                                            }
                                        }
                                    }
                                }
                                catch {
                                    Write-LogEntry -LogName $Log -LogEntryText "Error getting roles for site user/group $userLogin : $_" -LogLevel "ERROR"
                                }
                            }
                            
                            # Add users/groups with direct permissions (and meaningful roles for regular users)
                            if ($hasDirectPermissions -and ($isEveryone -or $userRoles.Count -gt 0)) {
                                # Create user object with roles
                                $userObject = [PSCustomObject]@{
                                    Name      = if ($isEveryone) { $userName }
                                    Email     = $userEmail
                                    LoginName = $userLogin
                                    Roles     = ($userRoles | Select-Object -Unique) -join ','
                                }
                                
                                # Add to the collection
                                $siteCollectionData[$siteUrl]["Site Level Users"].Add($userObject)
                                Write-LogEntry -LogName $Log -LogEntryText "Added site level principal: $($userObject.Name) with roles: $($userObject.Roles)" -LogLevel "DEBUG"
                            }
                            else {
                                Write-LogEntry -LogName $Log -LogEntryText "Skipping user/group $userName - No direct meaningful permissions" -LogLevel "DEBUG"
                            }
                        }
                        catch {
                            Write-LogEntry -LogName $Log -LogEntryText "Error processing site level user/group $($siteUser.Title): $_" -LogLevel "ERROR"
                        }
                    }
                    
                    Write-LogEntry -LogName $Log -LogEntryText "Completed processing site level users/groups for $siteUrl - Added $($siteCollectionData[$siteUrl]['Site Level Users'].Count) entries" -LogLevel "DEBUG"
                }
            }
            catch {
                Write-LogEntry -LogName $Log -LogEntryText "Error checking for site level users at site collection level for $siteUrl : $_" -LogLevel "ERROR"
            }
        }
        catch { Write-LogEntry -LogName $Log -LogEntryText "ERROR: Could not connect to site $siteUrl. Skipping SP Group/User processing. $_" -LogLevel "ERROR"; continue }

        # --- Version Policy Processing ---
        try {
            Write-LogEntry -LogName $Log -LogEntryText "Retrieving version policy for site $siteUrl" -LogLevel "DEBUG"
            
            $versionPolicy = Invoke-PnPWithRetry -ScriptBlock { 
                Get-PnPSiteVersionPolicy 
            } -Operation "Get-PnPSiteVersionPolicy for $siteUrl" -LogName $Log
            
            if ($versionPolicy) {
                Write-LogEntry -LogName $Log -LogEntryText "Successfully retrieved version policy for site $siteUrl" -LogLevel "DEBUG"
                
                # Debug output to verify the actual values
                Write-LogEntry -LogName $Log -LogEntryText "Version policy values - DefaultTrimMode: $($versionPolicy.DefaultTrimMode), DefaultExpireAfterDays: $($versionPolicy.DefaultExpireAfterDays), MajorVersionLimit: $($versionPolicy.MajorVersionLimit)" -LogLevel "DEBUG"
                
                # Update site data with version policy details - Pass values explicitly to avoid type conversion issues
                $expireDays = if ($null -eq $versionPolicy.DefaultExpireAfterDays) { -1 } else { [int]$versionPolicy.DefaultExpireAfterDays }
                $versionLimit = if ($null -eq $versionPolicy.MajorVersionLimit) { -1 } else { [int]$versionPolicy.MajorVersionLimit }
                
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops `
                    -DefaultTrimMode $versionPolicy.DefaultTrimMode `
                    -DefaultExpireAfterDays $expireDays `
                    -MajorVersionLimit $versionLimit
            }
            else {
                Write-LogEntry -LogName $Log -LogEntryText "Warning: No version policy found for site $siteUrl" -LogLevel "WARNING"
            }
        }
        catch {
            Write-LogEntry -LogName $Log -LogEntryText "Error retrieving version policy for site $siteUrl : $_" -LogLevel "ERROR"
        }

        # --- Site Collection Administrators Processing ---
        try {
            Write-LogEntry -LogName $Log -LogEntryText "Retrieving site collection administrators for site $siteUrl" -LogLevel "DEBUG"
            
            $siteAdmins = Invoke-PnPWithRetry -ScriptBlock { 
                Get-PnPSiteCollectionAdmin 
            } -Operation "Get-PnPSiteCollectionAdmin for $siteUrl" -LogName $Log

            if ($siteAdmins -and $siteAdmins.Count -gt 0) {
                Write-LogEntry -LogName $Log -LogEntryText "Found $($siteAdmins.Count) site collection administrators on $siteUrl" -LogLevel "DEBUG"
                
                foreach ($admin in $siteAdmins) {
                    if (!$admin -or !$admin.LoginName) { 
                        Write-LogEntry -LogName $Log -LogEntryText "Skipping null site admin $siteUrl" -LogLevel "WARNING"
                        continue 
                    }
                    
                    $adminName = $admin.Title
                    $adminEmail = $admin.Email
                    
                    # Get additional info from Azure AD if it's a user account
                    if ($admin.LoginName -like '*@*' -and $admin.PrincipalType -eq 'User') {
                        try {
                            $aadUser = Invoke-PnPWithRetry -ScriptBlock { 
                                Get-PnPAzureADUser -Identity $admin.LoginName -ErrorAction SilentlyContinue 
                            } -Operation "Get-PnPAzureADUser for admin $($admin.LoginName)" -LogName $Log
                            
                            if ($aadUser) { 
                                $adminName = $aadUser.DisplayName
                                $adminEmail = $aadUser.Mail
                            }
                        }
                        catch { 
                            Write-LogEntry -LogName $Log -LogEntryText "Warn: Getting AAD User info for admin '$($admin.LoginName)' failed: $_" -LogLevel "WARNING" 
                        }
                    }
                    
                    # Add the admin to the site collection data
                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -SiteAdminName $adminName -SiteAdminEmail $adminEmail
                }
            }
            else {
                Write-LogEntry -LogName $Log -LogEntryText "No site collection administrators found for $siteUrl or unable to retrieve them" -LogLevel "WARNING"
            }
        }
        catch {
            Write-LogEntry -LogName $Log -LogEntryText "Error retrieving site collection administrators for $siteUrl : $_" -LogLevel "ERROR"
        }

        # --- Microsoft 365 Group Processing (if applicable) ---
        if ($null -ne $siteprops.GroupId -and $siteprops.GroupId -ne [System.Guid]::Empty) {
            Write-LogEntry -LogName $Log -LogEntryText "Site $siteUrl connected M365 Group: $($siteprops.GroupId)." -LogLevel "DEBUG"
            try {
                # Get M365 Group Details
                $AADGroups = Invoke-PnPWithRetry -ScriptBlock { 
                    Get-PnPMicrosoft365Group -Identity $siteprops.GroupId 
                } -Operation "Get-PnPMicrosoft365Group for $($siteprops.GroupId)" -LogName $Log
                if ($AADGroups) {
                    Write-LogEntry -LogName $Log -LogEntryText "Successfully retrieved AAD Group details for $($siteprops.GroupId)." -LogLevel "DEBUG"
                    # Update site data with AAD Group details
                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -AADGroups $AADGroups
                }
                else {
                    Write-LogEntry -LogName $Log -LogEntryText "Warning: Get-PnPMicrosoft365Group returned null for Group ID $($siteprops.GroupId) on site $siteUrl." -LogLevel "WARNING"
                }

                # Get M365 Group Owners and Members
                $groupownersRaw = Invoke-PnPWithRetry -ScriptBlock { 
                    Get-PnPMicrosoft365GroupOwners -Identity $siteprops.GroupId 
                } -Operation "Get-PnPMicrosoft365GroupOwners for $($siteprops.GroupId)" -LogName $Log
                
                $groupmembersRaw = Invoke-PnPWithRetry -ScriptBlock { 
                    Get-PnPMicrosoft365GroupMembers -Identity $siteprops.GroupId 
                } -Operation "Get-PnPMicrosoft365GroupMembers for $($siteprops.GroupId)" -LogName $Log
                
                Write-LogEntry -LogName $Log -LogEntryText "Retrieved $($groupownersRaw.Count) owners / $($groupmembersRaw.Count) members for M365 Group $($siteprops.GroupId)" -LogLevel "DEBUG"

                # Process Owners & Members
                foreach ($owner in $groupownersRaw) {
                    try {
                        $aadOwnerUser = Invoke-PnPWithRetry -ScriptBlock { 
                            Get-PnPAzureADUser -Identity $owner.Id 
                        } -Operation "Get-PnPAzureADUser for owner $($owner.Id)" -LogName $Log
                        
                        if ($aadOwnerUser) { Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -EntraGroupOwner $aadOwnerUser.DisplayName -EntraGroupOwnerEmail $aadOwnerUser.Mail }
                        else { Write-LogEntry -LogName $Log -LogEntryText "Could not find AAD details M365 Owner ID: $($owner.Id)" -LogLevel "WARNING" }
                    }
                    catch { Write-LogEntry -LogName $Log -LogEntryText "Error getting AAD details M365 Owner ID $($owner.Id): $_" -LogLevel "ERROR" }
                }
                foreach ($member in $groupmembersRaw) {
                    try {
                        $aadMemberUser = Invoke-PnPWithRetry -ScriptBlock { 
                            Get-PnPAzureADUser -Identity $member.Id 
                        } -Operation "Get-PnPAzureADUser for member $($member.Id)" -LogName $Log
                        
                        if ($aadMemberUser) { Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -EntraGroupMember $aadMemberUser.DisplayName -EntraGroupMemberEmail $aadMemberUser.Mail }
                        else { Write-LogEntry -LogName $Log -LogEntryText "Could not find AAD details M365 Member ID: $($member.Id)" -LogLevel "WARNING" }
                    }
                    catch { Write-LogEntry -LogName $Log -LogEntryText "Error getting AAD details M365 Member ID $($member.Id): $_" -LogLevel "ERROR" }
                }
            }
            catch { Write-LogEntry -LogName $Log -LogEntryText "Warning: Could not retrieve M365 group info for $($siteprops.GroupId) site $siteUrl : $_" -LogLevel "WARNING" }
        }
        else { Write-LogEntry -LogName $Log -LogEntryText "Site $siteUrl not connected to M365 Group." -LogLevel "DEBUG" }

        # --- SharePoint Group Processing ---
        $spGroups = @()
        try {
            $spGroups = Invoke-PnPWithRetry -ScriptBlock { 
                Get-PnPGroup 
            } -Operation "Get-PnPGroup for $siteUrl" -LogName $Log
            
            Write-LogEntry -LogName $Log -LogEntryText "Found $($spGroups.Count) SP Groups on $siteUrl" -LogLevel "DEBUG"
        }
        catch { Write-LogEntry -LogName $Log -LogEntryText "Error retrieving SP groups for site $siteUrl : $_" -LogLevel "ERROR" }

        ForEach ($spGroup in $spGroups) {
            if (!$spGroup -or !$spGroup.Title) { Write-LogEntry -LogName $Log -LogEntryText "Skipping null SP group/title $siteUrl" -LogLevel "WARNING"; continue }

            $spGroupName = $spGroup.Title; $spGroupRolesString = ""
            Write-LogEntry -LogName $Log -LogEntryText "Processing SP Group: '$spGroupName' $siteUrl" -LogLevel "DEBUG"
            
            # Check if this is a sharing links group
            if ($spGroupName -like "SharingLinks*") {
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -SPGroupName $spGroupName
            }

            # Get SP Group Roles with throttling handling
            try {
                $web = Invoke-PnPWithRetry -ScriptBlock { 
                    Get-PnPWeb -Includes RoleAssignments 
                } -Operation "Get-PnPWeb with RoleAssignments for $siteUrl" -LogName $Log
                
                $groupRoleAssignments = $web.RoleAssignments
                if ($groupRoleAssignments) {
                    $rolesList = [System.Collections.Generic.List[string]]::new()
                    foreach ($roleAssignment in $groupRoleAssignments) {
                        $roleAssignmentWithDefs = Invoke-PnPWithRetry -ScriptBlock { 
                            Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings 
                        } -Operation "Get-PnPProperty RoleDefinitionBindings for group $spGroupName" -LogName $Log
                        
                        foreach ($roleDef in $roleAssignmentWithDefs) { 
                            if ($roleDef -and $roleDef.Name -and -not $rolesList.Contains($roleDef.Name)) { 
                                $rolesList.Add($roleDef.Name) 
                            } 
                        }
                    }
                    $spGroupRolesString = $rolesList -join ','
                }
                else { 
                    Write-LogEntry -LogName $Log -LogEntryText "No role assignments SP group '$spGroupName' $siteUrl" -LogLevel "DEBUG" 
                }
            }
            catch { 
                Write-LogEntry -LogName $Log -LogEntryText "Error retrieving roles SP group '$spGroupName' $siteUrl : $_" -LogLevel "ERROR" 
            }

            # Update site data with the Group Name and its Roles
            Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -SPGroupName $spGroupName -SPGroupRoles $spGroupRolesString

            # Get SP Group Members with throttling handling
            $spGroupMembers = @()
            try {
                if ($spGroup.Id) { 
                    $spGroupMembers = Invoke-PnPWithRetry -ScriptBlock { 
                        Get-PnPGroupMember -Identity $spGroup.Id 
                    } -Operation "Get-PnPGroupMember for group $spGroupName" -LogName $Log
                }
                else { 
                    Write-LogEntry -LogName $Log -LogEntryText "SP Group '$spGroupName' null ID." -LogLevel "WARNING" 
                }

                foreach ($member in $spGroupMembers) {
                    if (!$member -or !$member.LoginName) { 
                        Write-LogEntry -LogName $Log -LogEntryText "Skipping null/empty member SP group '$spGroupName'." -LogLevel "WARNING"
                        continue 
                    }

                    $spUserLogin = $member.LoginName
                    $spUserTitle = $member.Title
                    $spUserName = ""
                    $spUserEmail = ""
                    
                    # Check for spo-grid-all-users in the login name
                    if ($spUserLogin -like "*spo-grid-all-users*") {
                        Write-LogEntry -LogName $Log -LogEntryText "Found 'Everyone' user (spo-grid-all-users) in group '$spGroupName' on $siteUrl" -LogLevel "DEBUG"
                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -SPUserLoginName $spUserLogin
                    }

                    try {
                        $pnpUser = Invoke-PnPWithRetry -ScriptBlock { 
                            Get-PnPUser -Identity $spUserLogin -ErrorAction SilentlyContinue 
                        } -Operation "Get-PnPUser for $spUserLogin" -LogName $Log
                        
                        if ($pnpUser) {
                            $spUserName = $pnpUser.Title
                            $spUserEmail = $pnpUser.Email
                            
                            if ($pnpUser.LoginName -like '*@*' -and $pnpUser.PrincipalType -eq 'User') {
                                try {
                                    $aadUser = Invoke-PnPWithRetry -ScriptBlock { 
                                        Get-PnPAzureADUser -Identity $pnpUser.Email 
                                    } -Operation "Get-PnPAzureADUser for $($pnpUser.Email)" -LogName $Log
                                    
                                    if ($aadUser) { 
                                        $spUserName = $aadUser.DisplayName
                                        $spUserEmail = $aadUser.Mail
                                    }
                                    else { 
                                        Write-LogEntry -LogName $Log -LogEntryText "AAD User not found '$($pnpUser.LoginName)'." -LogLevel "DEBUG" 
                                    }
                                }
                                catch { 
                                    Write-LogEntry -LogName $Log -LogEntryText "Warn: Getting AAD User '$($pnpUser.LoginName)' failed: $_" -LogLevel "WARNING" 
                                }
                            }
                            elseif ($pnpUser.PrincipalType -ne 'User') { 
                                Write-LogEntry -LogName $Log -LogEntryText "Login '$spUserLogin' is $($pnpUser.PrincipalType)." -LogLevel "DEBUG"
                                $spUserName = if ($pnpUser.Title) { $pnpUser.Title } else { $spUserLogin } 
                            }
                        }
                        else { 
                            Write-LogEntry -LogName $Log -LogEntryText "Warn: Get-PnPUser failed '$spUserLogin'." -LogLevel "WARNING"
                            $spUserName = $spUserTitle 
                        }

                        # Call Update-SiteCollectionData for the specific user/group combo
                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -AssociatedSPGroup $spGroupName -SPUserName $spUserName -SPUserTitle $spUserTitle -SPUserEmail $spUserEmail -SPUserLoginName $spUserLogin
                    }
                    catch { 
                        Write-LogEntry -LogName $Log -LogEntryText "Error processing member '$($member.LoginName)' SP group '$spGroupName' $siteUrl : $_" -LogLevel "ERROR"
                        # Fallback
                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -AssociatedSPGroup $spGroupName -SPUserName $member.Title -SPUserTitle $member.Title 
                    }
                } # End foreach SP Group Member
            }
            catch { 
                Write-LogEntry -LogName $Log -LogEntryText "Error retrieving members SP group '$spGroupName' $siteUrl : $_" -LogLevel "ERROR" 
            }
        } # End foreach SP Group
    }
    catch {
        Write-LogEntry -LogName $Log -LogEntryText "FATAL Error main processing block $siteUrl : $_" -LogLevel "ERROR"
        continue # Continue to the next site
    }
    
    # Export this site's data to CSV immediately after processing
    Export-SiteCollectionToCSV -SiteUrl $siteUrl -CsvPath $outputfile
    Write-Host "Exported data for site $processedCount/$totalSites to CSV" -ForegroundColor Green

} # End foreach Site

# Disconnect
Disconnect-PnPOnline
Write-LogEntry -LogName $Log -LogEntryText "Disconnected from PnP Online. Script finished."
Write-Host "Script finished. Log file located at: $log" -ForegroundColor Green
Write-Host "Output CSV located at: $outputfile" -ForegroundColor Green
