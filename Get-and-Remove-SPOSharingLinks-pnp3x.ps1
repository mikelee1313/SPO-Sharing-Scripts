<#
.SYNOPSIS
    Identifies and processes SharePoint Online sharing links across the tenant.

.DESCRIPTION
    This script scans SharePoint Online sites to identify all sharing links, with a focus on Organization sharing links. 
    It can optionally convert Organization sharing links to direct permissions and clean up corrupted sharing groups.
    The script supports scanning all sites in a tenant or a specific list of sites from a CSV file.

.PARAMETER tenantname
    The name of your Microsoft 365 tenant (without .onmicrosoft.com).

.PARAMETER appID
    The Entra (Azure AD) application ID used for authentication.

.PARAMETER thumbprint
    The certificate thumbprint for authentication.

.PARAMETER tenant
    The tenant ID (GUID) for your Microsoft 365 tenant.

.PARAMETER searchRegion
    The region for Microsoft Graph search operations (e.g., "NAM", "EUR").

.PARAMETER convertOrganizationLinks
    When set to $true, the script converts Organization sharing links to direct permissions.
    When set to $false, the script only inventories sharing links without modifying them.

.PARAMETER cleanupCorruptedSharingGroups
    When set to $true, the script attempts to clean up empty or corrupted sharing groups.
    When set to $false, no cleanup of sharing groups is performed.
    Note: This is automatically set to $true when convertOrganizationLinks is $true (remediation mode).

.PARAMETER debugLogging
    When set to $true, the script logs detailed DEBUG operations for troubleshooting.
    When set to $false, only INFO and ERROR operations are logged.

.PARAMETER inputfile
    Optional. Path to a CSV file containing either:
    1. A simple list of SharePoint site URLs (one URL per line or with "URL" header)
    2. The output CSV from a previous run of this script in report mode
    
    When using the script's own CSV output as input, the script will:
    - Only process sites that have Organization sharing links (identified by group names containing "Organization")
    - Automatically set $convertOrganizationLinks to $true for remediation
    - Skip other types of sharing links for focused remediation
    
    If not specified, the script will process all sites in the tenant.

.OUTPUTS
    - CSV file containing detailed information about sharing links found
    - Log file with operation details and errors

.NOTES
    Authors: Mike Lee
    Date: 7/1/2025

    - Requires PnP.PowerShell module
    - Requires an Entra app registration with appropriate SharePoint permissions
    - Requires a certificate for authentication
    - The app must have Sites.FullControl.All and User.Read.All permissions
    - For optimal performance, use a certificate-based app rather than client secret

.Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

.EXAMPLE
    # Process all sites, only inventory sharing links without modifications (report mode)
    .\Get-and-Remove-SPOSharingLinks-pnp2x.ps1

.EXAMPLE
    # Process sites from a simple CSV file with URLs and convert Organization links to direct permissions
    $tenantname = "contoso"
    $appID = "12345678-1234-1234-1234-1234567890ab"
    $thumbprint = "1234567890ABCDEF1234567890ABCDEF12345678"
    $tenant = "12345678-1234-1234-1234-1234567890ab"
    $inputfile = "C:\temp\sitelist.csv"
    $convertOrganizationLinks = $true
    # Note: $cleanupCorruptedSharingGroups will automatically be set to $true in remediation mode
    .\Get-and-Remove-SPOSharingLinks-pnp2x.ps1

.EXAMPLE
    # Two-step process: Report then Remediate
    # Step 1: Run in report mode to generate CSV output
    $convertOrganizationLinks = $false
    .\Get-and-Remove-SPOSharingLinks-pnp2x.ps1
    
    # Step 2: Use the generated CSV to remediate only Organization links
    $inputfile = "C:\temp\SPO_SharingLinks_2025-07-01_14-30-15.csv"
    # Note: $convertOrganizationLinks will be automatically set to $true when using script's CSV output
    .\Get-and-Remove-SPOSharingLinks-pnp2x.ps1
#>


# ----------------------------------------------
# Set Variables
# ----------------------------------------------
$tenantname = "m365x61250205"                                   # This is your tenant name
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51"                # This is your Tenant ID
$searchRegion = "NAM"                                           # Region for Microsoft Graph search
$convertOrganizationLinks = $false                              # Set to $false ro report only, $true to convert Organization sharing links to direct permissions
$debugLogging = $true                                           # Set to $true for detailed DEBUG logging, $false for INFO and ERROR logging only

# ----------------------------------------------
# Initialize Parameters - Do not change
# ----------------------------------------------
$sites = @()
$cleanupCorruptedSharingGroups = $false                         # Set to $false to skip cleanup of empty/corrupted sharing groups (Note: automatically enabled in remediation mode)
# Auto-enable cleanup when in remediation mode (converting Organization links)
if ($convertOrganizationLinks) {
    $cleanupCorruptedSharingGroups = $true
    Write-InfoLog -LogName $Log -LogEntryText "Auto-enabled cleanup of corrupted sharing groups because remediation mode is active"
}
$inputfile = $null
$log = $null
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# ----------------------------------------------
# Input / Output and Log Files
# ----------------------------------------------
$inputfile = $null #comment this line to run against all SPO Sites, otherwise use an input file.
$log = "$env:TEMP\" + 'SPOSharingLinks' + $date + '_' + "logfile.log"
# Initialize sharing links output file
$sharingLinksOutputFile = "$env:TEMP\" + 'SPO_SharingLinks_' + $date + '.csv'

# ----------------------------------------------
# Logging Function
# ----------------------------------------------
Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $Level = "INFO" # INFO, DEBUG, ERROR
    )
    
    # Always log INFO and ERROR messages
    # Only log DEBUG messages when debug logging is enabled
    if ($Level -eq "ERROR" -or $Level -eq "INFO" -or ($Level -eq "DEBUG" -and $debugLogging)) {
        if ($LogName -ne $null) {
            # log the date and time in the text file along with the data passed
            "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) [$Level] : $LogEntryText" | Out-File -FilePath $LogName -append;
        }
    }
}

# ----------------------------------------------
# Logging Helper Functions
# ----------------------------------------------
Function Write-InfoLog {
    param(
        [string] $LogName,
        [string] $LogEntryText
    )
    # Always log INFO messages
    Write-LogEntry -LogName $LogName -LogEntryText $LogEntryText -Level "INFO"
}

Function Write-DebugLog {
    param(
        [string] $LogName,
        [string] $LogEntryText
    )
    # Only log DEBUG messages when debug logging is enabled
    Write-LogEntry -LogName $LogName -LogEntryText $LogEntryText -Level "DEBUG"
}

Function Write-ErrorLog {
    param(
        [string] $LogName,
        [string] $LogEntryText
    )
    # Always log ERROR messages
    Write-LogEntry -LogName $LogName -LogEntryText $LogEntryText -Level "ERROR"
}

# ----------------------------------------------
# Connection Parameters
# ----------------------------------------------
$connectionParams = @{
    ClientId      = $appID
    Thumbprint    = $thumbprint
    Tenant        = $tenant
    WarningAction = 'SilentlyContinue'
}

# ----------------------------------------------
# Throttling Handling Function
# ----------------------------------------------
Function Invoke-WithThrottleHandling {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock] $ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [int] $MaxRetries = 5,
        
        [Parameter(Mandatory = $false)]
        [string] $Operation = "SharePoint Operation"
    )
    
    $retryCount = 0
    $success = $false
    $result = $null
    
    while (-not $success -and $retryCount -le $MaxRetries) {
        try {
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            $errorMessage = $_.Exception.Message
            
            # Check if this is a throttling error
            $isThrottling = $false
            $waitTime = 10 # Default wait time in seconds
            
            # Check for common throttling status codes
            if ($null -ne $_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                
                if ($statusCode -eq 429 -or $statusCode -eq 503) {
                    $isThrottling = $true
                    
                    # Try to get the Retry-After header
                    $retryAfterHeader = $_.Exception.Response.Headers["Retry-After"]
                    
                    if ($retryAfterHeader) {
                        $waitTime = [int]$retryAfterHeader
                        Write-DebugLog -LogName $Log -LogEntryText "Throttling detected for $Operation. Retry-After header: $waitTime seconds."
                    }
                    else {
                        # Use exponential backoff if no Retry-After header
                        $waitTime = [Math]::Pow(2, $retryCount) * 10
                        Write-DebugLog -LogName $Log -LogEntryText "Throttling detected for $Operation. No Retry-After header. Using backoff: $waitTime seconds."
                    }
                }
            }
            # PnP specific throttling detection
            elseif ($errorMessage -match "throttl|Too many requests|429|503|Request limit exceeded") {
                $isThrottling = $true
                
                # Extract wait time from error message if available
                if ($errorMessage -match "Try again in (\d+) (seconds|minutes)") {
                    $timeValue = [int]$matches[1]
                    $timeUnit = $matches[2]
                    
                    $waitTime = if ($timeUnit -eq "minutes") { $timeValue * 60 } else { $timeValue }
                    Write-DebugLog -LogName $Log -LogEntryText "PnP throttling detected for $Operation. Waiting for $waitTime seconds."
                }
                else {
                    # Use exponential backoff
                    $waitTime = [Math]::Pow(2, $retryCount) * 10
                    Write-DebugLog -LogName $Log -LogEntryText "PnP throttling detected for $Operation. Using backoff: $waitTime seconds."
                }
            }
            
            if ($isThrottling) {
                $retryCount++
                
                if ($retryCount -le $MaxRetries) {
                    Write-Host "  Throttling detected for $Operation. Retrying in $waitTime seconds... (Attempt $retryCount of $MaxRetries)" -ForegroundColor Yellow
                    Write-DebugLog -LogName $Log -LogEntryText "Waiting $waitTime seconds before retry #$retryCount for $Operation."
                    Start-Sleep -Seconds $waitTime
                    continue
                }
            }
            
            # If we reach here, it's either not throttling or we've exceeded retries
            Write-Host "Error in $Operation (Retry #$retryCount): $errorMessage" -ForegroundColor Red
            Write-ErrorLog -LogName $Log -LogEntryText "Error in $Operation (Retry #$retryCount): $errorMessage"
            throw
        }
    }
    
    return $result
}

# ----------------------------------------------
# Connect to Admin Center initially
# ----------------------------------------------
try {
    $adminUrl = 'https://' + $tenantname + '-admin.sharepoint.com'
    Connect-PnPOnline -Url $adminUrl @connectionParams
    Write-InfoLog -LogName $Log -LogEntryText "Successfully connected to SharePoint Admin Center: $adminUrl"
}
catch {
    Write-Host "Error connecting to SharePoint Admin Center ($adminUrl): $_" -ForegroundColor Red
    Write-ErrorLog -LogName $Log -LogEntryText "Error connecting to SharePoint Admin Center ($adminUrl): $_"
    exit
}

# ----------------------------------------------
# Get Site List
# ----------------------------------------------
if ($inputfile -and (Test-Path -Path $inputfile)) {
    Write-Host "Processing input file: $inputfile" -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "Processing input file: $inputfile"
    
    try {
        # Check if this is the script's CSV output format or a simple URL list
        $firstLine = Get-Content -Path $inputfile -TotalCount 1
        
        if ($firstLine -and $firstLine.Contains("Sharing Group Name")) {
            # This is the script's CSV output format
            Write-Host "Detected script's CSV output format - will process Organization sharing links only" -ForegroundColor Cyan
            Write-InfoLog -LogName $Log -LogEntryText "Input file detected as script's CSV output format"
            
            # Import the full CSV and filter for Organization sharing links only
            $csvData = Import-Csv -Path $inputfile
            $organizationEntries = $csvData | Where-Object { 
                $_."Sharing Group Name" -like "*Organization*" -and
                -not [string]::IsNullOrWhiteSpace($_."Site URL")
            }
            
            if ($organizationEntries.Count -eq 0) {
                Write-Host "No Organization sharing links found in the input CSV file" -ForegroundColor Yellow
                Write-InfoLog -LogName $Log -LogEntryText "No Organization sharing links found in input CSV"
                $sites = @()
            }
            else {
                # Group by Site URL to get unique sites and force conversion mode
                $siteGroups = $organizationEntries | Group-Object "Site URL"
                $sites = $siteGroups | ForEach-Object { [PSCustomObject]@{ URL = $_.Name } }
                
                # Auto-enable conversion when using CSV output
                if (-not $convertOrganizationLinks) {
                    Write-Host "Auto-enabling Organization link conversion for CSV input mode" -ForegroundColor Green
                    $convertOrganizationLinks = $true
                    $cleanupCorruptedSharingGroups = $true
                    Write-InfoLog -LogName $Log -LogEntryText "Auto-enabled remediation mode and cleanup for CSV input containing Organization links"
                }
                
                Write-Host "Found $($sites.Count) sites with Organization sharing links for remediation" -ForegroundColor Green
                Write-InfoLog -LogName $Log -LogEntryText "Parsed $($sites.Count) sites with Organization sharing links from CSV input"
            }
        }
        else {
            # This is a simple site URL list
            Write-Host "Input file appears to be a simple site URL list" -ForegroundColor Yellow
            Write-InfoLog -LogName $Log -LogEntryText "Input file detected as simple site URL list"
            $sites = Import-csv -path $inputfile -Header 'URL'
        }
    }
    catch {
        Write-Host "Error reading input file '$inputfile': $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error reading input file '$inputfile': $_"
        exit
    }
}
else {
    Write-Host "Getting site list from tenant (this might take a while)..." -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "Getting sites using Get-PnPTenantSite (no input file specified or found)"
    try {
        # Get sites with optimized filtering to reduce memory usage and improve performance
        $sites = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPTenantSite -IncludeOneDriveSites:$false | Where-Object {
                $_.Template -notmatch "SRCHCEN|MYSITE|APPCATALOG|PWS|POINTPUBLISHINGTOPIC|SPSMSITEHOST|EHS|REVIEWCTR|TENANTADMIN" -and
                $_.Status -eq "Active" -and
                -not [string]::IsNullOrEmpty($_.Url)
            }
        } -Operation "Get-PnPTenantSite with optimized filtering"
        
        Write-Host "Found $($sites.Count) sites for processing after filtering." -ForegroundColor Green
        Write-InfoLog -LogName $log -LogEntryText "Retrieved and filtered to $($sites.Count) sites for processing."
    }
    catch {
        Write-Host "Error getting site list from tenant: $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error getting site list from tenant: $_"
        exit
    }
}

# ----------------------------------------------
# Initialize a hashtable to store site collection data (keyed by URL)
# ----------------------------------------------
$siteCollectionData = @{}

# ----------------------------------------------
# Initialize the sharing links output file with headers
# ----------------------------------------------
$sharingLinksHeaders = "Site URL,Site Owner,IB Mode,IB Segment,Site Template,Sharing Group Name,Sharing Link Members,File URL,File Owner,IsTeamsConnected,SharingCapability,Last Content Modified,Link Removed"
Set-Content -Path $sharingLinksOutputFile -Value $sharingLinksHeaders
Write-InfoLog -LogName $Log -LogEntryText "Initialized sharing links output file: $sharingLinksOutputFile"

# ----------------------------------------------
# Function to handle consolidated site data
# ----------------------------------------------
Function Update-SiteCollectionData {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [object] $SiteProperties,
        [string] $SPGroupName = "",
        # --- Parameters for SP User ---
        [string] $AssociatedSPGroup = "",
        [string] $SPUserName = "",
        [string] $SPUserTitle = "",
        [string] $SPUserEmail = ""
    )

    # Create site entry if it doesn't exist
    if (-not $siteCollectionData.ContainsKey($SiteUrl)) {
        $siteCollectionData[$SiteUrl] = @{
            "URL"                     = $SiteProperties.Url
            "Owner"                   = $SiteProperties.Owner
            "IB Mode"                 = ($SiteProperties.InformationBarrierMode -join ',')
            "IB Segment"              = ($SiteProperties.InformationBarrierSegments -join ',')
            "Template"                = $SiteProperties.Template
            "SharingCapability"       = $SiteProperties.SharingCapability
            "IsTeamsConnected"        = $SiteProperties.IsTeamsConnected
            "LastContentModifiedDate" = $SiteProperties.LastContentModifiedDate
            # Site-specific lists
            "SP Groups On Site"       = [System.Collections.Generic.List[string]]::new()
            "SP Users"                = [System.Collections.Generic.List[PSObject]]::new()
            "Has Sharing Links"       = $false # Property to track if sharing links are being used
            "Link Removal Status"     = @{} # Track which sharing groups had their links removed
        }
    }

    # Check for SharingLinks groups
    if (-not [string]::IsNullOrWhiteSpace($SPGroupName) -and $SPGroupName -like "SharingLinks*") {
        $siteCollectionData[$SiteUrl]["Has Sharing Links"] = $true
        
        # Initialize link removal status to False for all sharing groups by default
        if (-not $siteCollectionData[$SiteUrl]["Link Removal Status"].ContainsKey($SPGroupName)) {
            $siteCollectionData[$SiteUrl]["Link Removal Status"][$SPGroupName] = $false
        }
    }

    # Add SP Group if provided and not already present for this site
    if (-not [string]::IsNullOrWhiteSpace($SPGroupName)) {
        if (-not $siteCollectionData[$SiteUrl]["SP Groups On Site"].Contains($SPGroupName)) {
            $siteCollectionData[$SiteUrl]["SP Groups On Site"].Add($SPGroupName)
        }
    }

    # Add SharePoint User information (associated with a specific SP group)
    if (-not [string]::IsNullOrWhiteSpace($SPUserName) -and -not [string]::IsNullOrWhiteSpace($AssociatedSPGroup)) {
        $userObject = [PSCustomObject]@{
            AssociatedSPGroup = $AssociatedSPGroup # Store the group name
            Name              = $SPUserName
            Title             = $SPUserTitle
            Email             = $SPUserEmail
        }
        $siteCollectionData[$SiteUrl]["SP Users"].Add($userObject)
    }
}

# ----------------------------------------------
# Function to update link removal status for a sharing group
# ----------------------------------------------
Function Update-LinkRemovalStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [Parameter(Mandatory = $true)]
        [string] $SharingGroupName,
        [Parameter(Mandatory = $true)]
        [bool] $WasRemoved
    )
    
    if ($siteCollectionData.ContainsKey($SiteUrl)) {
        $siteCollectionData[$SiteUrl]["Link Removal Status"][$SharingGroupName] = $WasRemoved
        Write-DebugLog -LogName $Log -LogEntryText "Updated link removal status for $SharingGroupName on $SiteUrl : $WasRemoved"
    }
}

# ----------------------------------------------
# Function to process and write sharing links for a site
# ----------------------------------------------
Function Write-SiteSharingLinks {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [object] $SiteData
    )
    
    # Check if this site has sharing links groups
    $sharingLinkGroups = $SiteData."SP Groups On Site" | Where-Object { $_ -like "SharingLinks*" }
    
    if ($sharingLinkGroups.Count -gt 0) {
        Write-Host "  Processing $($sharingLinkGroups.Count) sharing link groups for site: $SiteUrl" -ForegroundColor Yellow
        Write-InfoLog -LogName $Log -LogEntryText "Processing $($sharingLinkGroups.Count) sharing link groups for site: $SiteUrl"
        
        foreach ($sharingGroup in $sharingLinkGroups) {
            # Get users in this sharing links group
            $groupMembers = $SiteData."SP Users" | Where-Object { $_.AssociatedSPGroup -eq $sharingGroup }
            
            if ($groupMembers.Count -gt 0) {
                # Format members as "Name <Email>"
                $membersFormatted = ($groupMembers | ForEach-Object {
                        $emailStr = if ($_.Email) { $_.Email | Out-String -NoNewline } else { "" }
                        "$($_.Name) <$emailStr>"
                    }) -join ';'
                
                # Get document details if available
                $documentUrl = "Not found"
                $documentOwner = "Not found"
                if ($SiteData.ContainsKey("DocumentDetails") -and $SiteData["DocumentDetails"].ContainsKey($sharingGroup)) {
                    $documentUrl = $SiteData["DocumentDetails"][$sharingGroup]["DocumentUrl"]
                    $documentOwner = $SiteData["DocumentDetails"][$sharingGroup]["DocumentOwner"]
                }
                
                # Get link removal status
                $linkRemoved = "False"
                if ($SiteData.ContainsKey("Link Removal Status") -and $SiteData["Link Removal Status"].ContainsKey($sharingGroup)) {
                    $linkRemoved = if ($SiteData["Link Removal Status"][$sharingGroup]) { "True" } else { "False" }
                }
                
                # Create CSV line
                $csvLine = [PSCustomObject]@{
                    "Site URL"              = $SiteData.URL
                    "Site Owner"            = $SiteData.Owner
                    "IB Mode"               = $SiteData."IB Mode"
                    "IB Segment"            = $SiteData."IB Segment"
                    "Site Template"         = $SiteData.Template
                    "Sharing Group Name"    = $sharingGroup
                    "Sharing Link Members"  = $membersFormatted
                    "File URL"              = $documentUrl
                    "File Owner"            = $documentOwner
                    "IsTeamsConnected"      = $SiteData.IsTeamsConnected
                    "SharingCapability"     = $SiteData.SharingCapability
                    "Last Content Modified" = $SiteData.LastContentModifiedDate
                    "Link Removed"          = $linkRemoved
                }
                
                # Write directly to the CSV file
                $csvLine | Export-Csv -Path $sharingLinksOutputFile -Append -NoTypeInformation -Force
                Write-DebugLog -LogName $Log -LogEntryText "  Wrote sharing link data for group: $sharingGroup"
            }
        }
    }
}

# ----------------------------------------------
# Function to convert Organization sharing links to direct permissions
# ----------------------------------------------
Function Convert-OrganizationSharingLinks {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl
    )
    
    Write-Host "  Checking for Organization sharing links on site: $SiteUrl" -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "Checking for Organization sharing links on site: $SiteUrl"
    
    try {
        # Connect to the specific site
        Connect-PnPOnline -Url $SiteUrl @connectionParams -ErrorAction Stop
        
        # First, clean up any corrupted sharing groups if enabled
        if ($cleanupCorruptedSharingGroups) {
            Remove-CorruptedSharingGroups -SiteUrl $SiteUrl
        }
        
        # Get all SharePoint groups that contain "Organization" in the name
        $organizationGroups = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPGroup | Where-Object { $_.Title -like "*Organization*" }
        } -Operation "Get Organization groups for $SiteUrl"
        
        if ($organizationGroups.Count -eq 0) {
            Write-DebugLog -LogName $Log -LogEntryText "No Organization sharing groups found on site: $SiteUrl"
            return
        }
        
        Write-Host "    Found $($organizationGroups.Count) Organization sharing groups" -ForegroundColor Green
        Write-InfoLog -LogName $Log -LogEntryText "Found $($organizationGroups.Count) Organization sharing groups on site: $SiteUrl"
        
        foreach ($orgGroup in $organizationGroups) {
            $groupName = $orgGroup.Title
            Write-Host "    Processing Organization group: $groupName" -ForegroundColor Cyan
            Write-DebugLog -LogName $Log -LogEntryText "Processing Organization group: $groupName"
            
            # Determine permission level based on group name
            $permissionLevel = ""
            if ($groupName -like "*OrganizationEdit*") {
                $permissionLevel = "Edit"
            }
            elseif ($groupName -like "*OrganizationView*") {
                $permissionLevel = "Read"
            }
            else {
                Write-ErrorLog -LogName $Log -LogEntryText "Warning: Could not determine permission level for group $groupName. Skipping."
                continue
            }
            
            # Get group members
            $groupMembers = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPGroupMember -Identity $orgGroup.Id
            } -Operation "Get members for Organization group $groupName"
            
            if ($groupMembers.Count -eq 0) {
                Write-DebugLog -LogName $Log -LogEntryText "No members found in Organization group: $groupName"
                continue
            }
            
            Write-Host "      Found $($groupMembers.Count) members in group $groupName" -ForegroundColor Green
            Write-InfoLog -LogName $Log -LogEntryText "Found $($groupMembers.Count) members in Organization group: $groupName"
            
            # Extract document information from group name
            $documentUrl = ""
            $documentId = ""
            
            if ($groupName -match "SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.") {
                $documentId = $matches[1]
                Write-DebugLog -LogName $Log -LogEntryText "Extracted document ID: $documentId from group: $groupName"
                
                # Try to find the document using Microsoft Graph
                try {
                    $graphToken = Invoke-WithThrottleHandling -ScriptBlock {
                        Get-PnPAccessToken
                    } -Operation "Get-PnPAccessToken for document search"
                    
                    if ($graphToken) {
                        $headers = @{
                            "Authorization" = "Bearer $graphToken"
                            "Content-Type"  = "application/json"
                        }
                        
                        $searchQuery = @{
                            requests = @(
                                @{
                                    entityTypes               = @("driveItem")
                                    query                     = @{
                                        queryString = "UniqueID:$documentId"
                                    }
                                    from                      = 0
                                    size                      = 25
                                    sharePointOneDriveOptions = @{
                                        includeContent = "sharedContent,privateContent"
                                    }
                                    region                    = $searchRegion
                                }
                            )
                        }
                        
                        $searchBody = $searchQuery | ConvertTo-Json -Depth 5
                        $searchUrl = "https://graph.microsoft.com/v1.0/search/query"
                        
                        $searchResults = Invoke-WithThrottleHandling -ScriptBlock {
                            Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body $searchBody
                        } -Operation "Microsoft Graph Search for document ID $documentId"
                        
                        if ($searchResults.value -and 
                            $searchResults.value[0].hitsContainers -and 
                            $searchResults.value[0].hitsContainers[0].hits -and 
                            $searchResults.value[0].hitsContainers[0].hits.Count -gt 0) {
                            
                            $hit = $searchResults.value[0].hitsContainers[0].hits[0]
                            $resource = $hit.resource
                            
                            if ($resource -and $resource.webUrl) {
                                $documentUrl = $resource.webUrl
                                Write-DebugLog -LogName $Log -LogEntryText "Found document URL: $documentUrl"
                            }
                            else {
                                Write-DebugLog -LogName $Log -LogEntryText "Document found but no webUrl property available"
                            }
                        }
                        else {
                            Write-DebugLog -LogName $Log -LogEntryText "No search results found for document ID: $documentId"
                        }
                    }
                }
                catch {
                    Write-DebugLog -LogName $Log -LogEntryText "Error searching for document via Graph API: $_"
                }
            }
            else {
                Write-DebugLog -LogName $Log -LogEntryText "Could not extract document ID from group name: $groupName"
            }
            
            # Log the final document URL status
            if ([string]::IsNullOrWhiteSpace($documentUrl)) {
                Write-DebugLog -LogName $Log -LogEntryText "Document URL is empty for group: $groupName - will use site-level permissions"
            }
            else {
                Write-DebugLog -LogName $Log -LogEntryText "Will attempt document-level permissions for: $documentUrl"
            }
            
            # Process each member
            foreach ($member in $groupMembers) {
                if (!$member -or !$member.LoginName) { continue }
                
                try {
                    Write-Host "        Processing member: $($member.Title)" -ForegroundColor White
                    Write-DebugLog -LogName $Log -LogEntryText "Processing member: $($member.Title) ($($member.LoginName))"
                    
                    # Remove user from the sharing group
                    $memberRemovalSuccess = $false
                    try {
                        Invoke-WithThrottleHandling -ScriptBlock {
                            # Try Force parameter first, fallback to no confirmation parameter
                            try {
                                Remove-PnPGroupMember -Identity $orgGroup.Id -LoginName $member.LoginName -Force
                                Write-DebugLog -LogName $Log -LogEntryText "Successfully removed $($member.LoginName) using Force parameter"
                            }
                            catch {
                                # Fallback if Force parameter is not supported
                                Remove-PnPGroupMember -Identity $orgGroup.Id -LoginName $member.LoginName
                                Write-DebugLog -LogName $Log -LogEntryText "Successfully removed $($member.LoginName) using fallback method"
                            }
                        } -Operation "Remove member $($member.LoginName) from group $groupName"
                        
                        $memberRemovalSuccess = $true
                        Write-Host "          Removed from sharing group: $groupName" -ForegroundColor Yellow
                        Write-InfoLog -LogName $Log -LogEntryText "Removed $($member.LoginName) from sharing group: $groupName"
                    }
                    catch {
                        Write-Host "          Error: Failed to remove member from sharing group: $_" -ForegroundColor Red
                        Write-ErrorLog -LogName $Log -LogEntryText "Error: Failed to remove $($member.LoginName) from sharing group $groupName : $_"
                        
                        # Continue processing - still try to grant permissions even if removal failed
                        Write-DebugLog -LogName $Log -LogEntryText "Continuing to process permissions for $($member.LoginName) despite removal failure"
                    }
                    
                    # Grant direct permissions to the document if we found it
                    if (-not [string]::IsNullOrWhiteSpace($documentUrl)) {
                        try {
                            # Parse the document URL to get the relative URL
                            $uri = [System.Uri]$documentUrl
                            $relativePath = $uri.AbsolutePath
                            
                            # Validate that we have a non-empty relative path
                            if ([string]::IsNullOrWhiteSpace($relativePath)) {
                                Write-DebugLog -LogName $Log -LogEntryText "Warning: Empty relative path from document URL: $documentUrl"
                                throw "Invalid document URL - empty relative path"
                            }
                            
                            Write-DebugLog -LogName $Log -LogEntryText "Attempting to grant permissions for document at: $relativePath"
                            
                            # Try CSOM method first (since it's working)
                            try {
                                Invoke-WithThrottleHandling -ScriptBlock {
                                    # Get the file and break inheritance
                                    $file = Get-PnPFile -Url $relativePath
                                    $listItem = $file.ListItemAllFields
                                    $listItem.BreakRoleInheritance($false, $true)
                                    $listItem.Context.Load($listItem)
                                    $listItem.Context.ExecuteQuery()
                                    
                                    # Get the role definition and user
                                    $roleDefinition = Get-PnPRoleDefinition -Identity $permissionLevel
                                    $user = Get-PnPUser -Identity $member.LoginName
                                    
                                    # Create role definition binding collection
                                    $roleBindings = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($listItem.Context)
                                    $roleBindings.Add($roleDefinition)
                                    
                                    # Create role assignment
                                    $roleAssignment = $listItem.RoleAssignments.Add($user, $roleBindings)
                                    $listItem.Context.Load($roleAssignment)
                                    $listItem.Context.ExecuteQuery()
                                } -Operation "Grant $permissionLevel permission using CSOM with proper binding collection"
                                
                                Write-Host "          Granted direct $permissionLevel permission to document (CSOM method)" -ForegroundColor Green
                                Write-InfoLog -LogName $Log -LogEntryText "Granted direct $permissionLevel permission to $($member.LoginName) for document using CSOM: $documentUrl"
                            }
                            catch {
                                $csomLastError = $_.Exception.Message
                                Write-DebugLog -LogName $Log -LogEntryText "CSOM method failed: $csomLastError"
                                # Fallback 1: Try to get the file and set permissions with better list handling
                                try {
                                    Write-DebugLog -LogName $Log -LogEntryText "Attempting PnP method for $relativePath"
                                    
                                    # Get the file with list information
                                    $file = Invoke-WithThrottleHandling -ScriptBlock {
                                        Get-PnPFile -Url $relativePath -AsListItem -ErrorAction SilentlyContinue
                                    } -Operation "Get file as list item for $relativePath"
                                    
                                    # Try multiple ways to get the list
                                    $targetList = $null
                                    
                                    if ($file -and $file.ParentList -and $file.ParentList.Id) {
                                        $targetList = $file.ParentList
                                        Write-DebugLog -LogName $Log -LogEntryText "Found parent list from file: $($targetList.Title)"
                                    }
                                    elseif ($file) {
                                        Write-DebugLog -LogName $Log -LogEntryText "File found but ParentList is null, trying to resolve list manually"
                                        
                                        # Try to get list from the file path
                                        $pathParts = $relativePath.TrimStart('/').Split('/')
                                        
                                        # Try different potential list names from the path
                                        $potentialListNames = @()
                                        if ($pathParts.Length -gt 1) {
                                            $potentialListNames += $pathParts[1] # Second part (after site)
                                        }
                                        if ($pathParts.Length -gt 2) {
                                            $potentialListNames += $pathParts[2] # Third part
                                        }
                                        
                                        # Add common library names
                                        $potentialListNames += @("Documents", "Shared Documents", "Site Assets")
                                        
                                        foreach ($listName in $potentialListNames) {
                                            try {
                                                $targetList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
                                                if ($targetList) {
                                                    Write-DebugLog -LogName $Log -LogEntryText "Successfully resolved list: $($targetList.Title)"
                                                    break
                                                }
                                            }
                                            catch {
                                                continue
                                            }
                                        }
                                    }
                                    else {
                                        Write-DebugLog -LogName $Log -LogEntryText "Could not retrieve file as list item"
                                    }
                                    
                                    if ($targetList -and $file -and $file.Id) {
                                        # Grant direct permissions to the file
                                        Invoke-WithThrottleHandling -ScriptBlock {
                                            Set-PnPListItemPermission -List $targetList.Id -Identity $file.Id -User $member.LoginName -AddRole $permissionLevel
                                        } -Operation "Grant $permissionLevel permission to $($member.LoginName) for document"
                                        
                                        Write-Host "          Granted direct $permissionLevel permission to document (PnP method)" -ForegroundColor Green
                                        Write-InfoLog -LogName $Log -LogEntryText "Granted direct $permissionLevel permission to $($member.LoginName) for document: $documentUrl"
                                    }
                                    else {
                                        $errorDetails = @()
                                        if (-not $targetList) { $errorDetails += "no target list" }
                                        if (-not $file) { $errorDetails += "no file" }
                                        elseif (-not $file.Id) { $errorDetails += "file has no ID" }
                                        throw "Could not proceed with PnP method: $($errorDetails -join ', ')"
                                    }
                                }
                                catch {
                                    $pnpLastError = $_.Exception.Message
                                    Write-DebugLog -LogName $Log -LogEntryText "PnP method failed: $pnpLastError"
                                    
                                    # Fallback 2: List search approach
                                    try {
                                        Write-DebugLog -LogName $Log -LogEntryText "Attempting list search method for $relativePath"
                                        
                                        Invoke-WithThrottleHandling -ScriptBlock {
                                            # Try to get the list from the file URL and use list-level permissions
                                            $pathParts = $relativePath.TrimStart('/').Split('/')
                                            $possibleListNames = @()
                                            
                                            # Try to extract list name from path (more comprehensive approach)
                                            if ($pathParts.Length -gt 1) {
                                                $possibleListNames += $pathParts[1] # Usually the library name
                                            }
                                            if ($pathParts.Length -gt 2) {
                                                $possibleListNames += $pathParts[2] # Alternative position
                                            }
                                            
                                            # Try common document library names
                                            $possibleListNames += @("Documents", "Shared Documents", "Site Assets", "SiteAssets")
                                            
                                            $permissionGranted = $false
                                            $fileName = [System.IO.Path]::GetFileName($relativePath)
                                            
                                            Write-DebugLog -LogName $Log -LogEntryText "Searching for file '$fileName' in lists: $($possibleListNames -join ', ')"
                                            
                                            foreach ($listName in $possibleListNames) {
                                                try {
                                                    $list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
                                                    if ($list) {
                                                        Write-DebugLog -LogName $Log -LogEntryText "Found list: $($list.Title), searching for file"
                                                        
                                                        # Try multiple approaches to find the file
                                                        $listItems = $null
                                                        
                                                        # Approach 1: Search by file name
                                                        try {
                                                            $listItems = Get-PnPListItem -List $list -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>$fileName</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
                                                        }
                                                        catch {
                                                            Write-DebugLog -LogName $Log -LogEntryText "File search by name failed in $($list.Title): $_"
                                                        }
                                                        
                                                        # Approach 2: Try without the XML query if that failed
                                                        if (-not $listItems -or $listItems.Count -eq 0) {
                                                            try {
                                                                $allItems = Get-PnPListItem -List $list -ErrorAction SilentlyContinue
                                                                $listItems = $allItems | Where-Object { $_.FieldValues.FileLeafRef -eq $fileName }
                                                            }
                                                            catch {
                                                                Write-DebugLog -LogName $Log -LogEntryText "Alternative file search failed in $($list.Title): $_"
                                                            }
                                                        }
                                                        
                                                        if ($listItems -and $listItems.Count -gt 0) {
                                                            $listItem = if ($listItems -is [array]) { $listItems[0] } else { $listItems }
                                                            Set-PnPListItemPermission -List $list.Id -Identity $listItem.Id -User $member.LoginName -AddRole $permissionLevel
                                                            $permissionGranted = $true
                                                            Write-InfoLog -LogName $Log -LogEntryText "Found and granted permissions to file in list: $($list.Title)"
                                                            break
                                                        }
                                                        else {
                                                            Write-DebugLog -LogName $Log -LogEntryText "File not found in list: $($list.Title)"
                                                        }
                                                    }
                                                    else {
                                                        Write-DebugLog -LogName $Log -LogEntryText "List not found: $listName"
                                                    }
                                                }
                                                catch {
                                                    Write-DebugLog -LogName $Log -LogEntryText "Error processing list '$listName': $_"
                                                    continue
                                                }
                                            }
                                            
                                            if (-not $permissionGranted) {
                                                throw "Could not locate the file '$fileName' in any of the searched document libraries: $($possibleListNames -join ', ')"
                                            }
                                        } -Operation "Grant $permissionLevel permission using list search approach"
                                        
                                        Write-Host "          Granted direct $permissionLevel permission to document (list search method)" -ForegroundColor Green
                                        Write-InfoLog -LogName $Log -LogEntryText "Granted direct $permissionLevel permission to $($member.LoginName) for document using list search: $documentUrl"
                                    }
                                    catch {
                                        $lastError = $_.Exception.Message
                                        Write-DebugLog -LogName $Log -LogEntryText "List search method failed: $lastError"
                                        throw "All document-level permission approaches failed. CSOM: $csomLastError, PnP: $pnpLastError, List search: $lastError"
                                    }
                                }
                            }
                        }
                        catch {
                            $fullErrorMessage = $_.Exception.Message
                            Write-Host "          Warning: Could not grant direct permissions to document: $fullErrorMessage" -ForegroundColor Red
                            Write-ErrorLog -LogName $Log -LogEntryText "Warning: Could not grant direct permissions to $($member.LoginName) for document $documentUrl : $fullErrorMessage"
                            
                            # Fallback: Grant permissions at site level
                            try {
                                Invoke-WithThrottleHandling -ScriptBlock {
                                    Set-PnPWebPermission -User $member.LoginName -AddRole $permissionLevel
                                } -Operation "Grant $permissionLevel permission to $($member.LoginName) at site level"
                                
                                Write-Host "          Granted $permissionLevel permission at site level (fallback)" -ForegroundColor Yellow
                                Write-InfoLog -LogName $Log -LogEntryText "Granted $permissionLevel permission to $($member.LoginName) at site level as fallback"
                            }
                            catch {
                                Write-Host "          Error: Could not grant any permissions: $_" -ForegroundColor Red
                                Write-ErrorLog -LogName $Log -LogEntryText "Error: Could not grant any permissions to $($member.LoginName): $_"
                            }
                        }
                    }
                    else {
                        # If we couldn't find the specific document, grant permissions at site level
                        Write-DebugLog -LogName $Log -LogEntryText "Document URL not found for group $groupName, granting site-level permissions"
                        try {
                            Invoke-WithThrottleHandling -ScriptBlock {
                                Set-PnPWebPermission -User $member.LoginName -AddRole $permissionLevel
                            } -Operation "Grant $permissionLevel permission to $($member.LoginName) at site level"
                            
                            Write-Host "          Granted $permissionLevel permission at site level (document not found)" -ForegroundColor Green
                            Write-InfoLog -LogName $Log -LogEntryText "Granted $permissionLevel permission to $($member.LoginName) at site level (document not found)"
                        }
                        catch {
                            Write-Host "          Error: Could not grant site-level permissions: $_" -ForegroundColor Red
                            Write-ErrorLog -LogName $Log -LogEntryText "Error: Could not grant site-level permissions to $($member.LoginName): $_"
                        }
                    }
                }
                catch {
                    Write-Host "        Error processing member $($member.Title): $_" -ForegroundColor Red
                    Write-ErrorLog -LogName $Log -LogEntryText "Error processing member $($member.Title) ($($member.LoginName)): $_"
                }
            }
            
            # Properly remove the sharing link using UnshareLink method
            Write-Host "      Removing Organization sharing link: $groupName" -ForegroundColor Yellow
            Write-DebugLog -LogName $Log -LogEntryText "Attempting to properly remove Organization sharing link: $groupName"
            
            try {
                # First verify the group is empty before removing
                $remainingMembers = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPGroupMember -Identity $orgGroup.Id -ErrorAction SilentlyContinue
                } -Operation "Check remaining members in Organization group $groupName"
                
                if ($remainingMembers -and $remainingMembers.Count -gt 0) {
                    Write-LogEntry -LogName $Log -LogEntryText "Warning: Group $groupName still has $($remainingMembers.Count) members, will not remove" -Level "INFO"
                    Write-Host "      Warning: Group still has members, skipping removal" -ForegroundColor Yellow
                }
                else {
                    # Group is empty, safe to remove the sharing link properly
                    $sharingLinkRemoved = $false
                    
                    # Try to remove using UnshareLink if we have document details
                    if ($documentUrl -and $documentId) {
                        try {
                            Write-DebugLog -LogName $Log -LogEntryText "Attempting to unshare link using PnP PowerShell methods"
                            
                            $result = Invoke-WithThrottleHandling -ScriptBlock {
                                # Parse document URL to get relative path
                                $uri = [System.Uri]$documentUrl
                                $relativePath = $uri.AbsolutePath
                                $linkRemoved = $false
                                
                                # Try PnP's Get-PnPFileSharingLink and Remove-PnPFileSharingLink first
                                try {
                                    $file = Get-PnPFile -Url $relativePath
                                    
                                    # Get all sharing links for this file using the correct parameter
                                    $sharingLinks = Get-PnPFileSharingLink -Identity $relativePath
                                    
                                    foreach ($sharingLink in $sharingLinks) {
                                        # Try to match the sharing link with our group
                                        if ($sharingLink.Id -and $groupName -like "*$($sharingLink.Id)*") {
                                            Write-LogEntry -LogName $Log -LogEntryText "Found matching sharing link with ID: $($sharingLink.Id)" -Level "INFO"
                                            
                                            # Remove the sharing link using the file URL and sharing link ID
                                            Remove-PnPFileSharingLink -FileUrl $relativePath -Id $sharingLink.Id -Force
                                            
                                            Write-Host "        Successfully removed sharing link using PnP methods" -ForegroundColor Green
                                            Write-InfoLog -LogName $Log -LogEntryText "Successfully removed sharing link with ID: $($sharingLink.Id)"
                                            $linkRemoved = $true
                                            break
                                        }
                                    }
                                }
                                catch {
                                    Write-LogEntry -LogName $Log -LogEntryText "PnP sharing link methods failed: $_" -Level "DEBUG"
                                    
                                    # Try alternative PnP approach - remove all sharing links for the file
                                    try {
                                        Write-DebugLog -LogName $Log -LogEntryText "Attempting to remove all sharing links for the file"
                                        
                                        # Use Set-PnPFileSharing -DisableSharing
                                        Set-PnPFileSharing -Url $relativePath -RemoveSharing
                                        
                                        Write-Host "        Successfully removed sharing using Set-PnPFileSharing" -ForegroundColor Green
                                        Write-InfoLog -LogName $Log -LogEntryText "Successfully removed sharing using Set-PnPFileSharing"
                                        $linkRemoved = $true
                                    }
                                    catch {
                                        Write-LogEntry -LogName $Log -LogEntryText "Set-PnPFileSharing also failed: $_" -Level "DEBUG"
                                        
                                        # Try CSOM approach with proper SharePoint Client methods
                                        try {
                                            Write-DebugLog -LogName $Log -LogEntryText "Attempting CSOM approach with SharePoint Client methods"
                                            
                                            # Get the web and file context
                                            $web = Get-PnPWeb
                                            $ctx = $web.Context
                                            
                                            # Get the file
                                            $file = $web.GetFileByServerRelativeUrl($relativePath)
                                            $ctx.Load($file)
                                            $ctx.ExecuteQuery()
                                            
                                            # Get the list item
                                            $listItem = $file.ListItemAllFields
                                            $ctx.Load($listItem)
                                            $ctx.ExecuteQuery()
                                            
                                            # Load Microsoft.SharePoint.Client.Sharing namespace
                                            Add-Type -Path (Get-Module PnP.PowerShell | Select-Object -ExpandProperty ModuleBase | Join-Path -ChildPath "Microsoft.SharePoint.Client.dll")
                                            
                                            # Try to get sharing information and remove it
                                            try {
                                                # Use ObjectSharingInformation to manage sharing
                                                $sharingInfo = [Microsoft.SharePoint.Client.Sharing.WebSharingManager]::GetObjectSharingInformation($ctx, $listItem, $false, $false, $false, $true, $true, $true, $true)
                                                $ctx.Load($sharingInfo)
                                                $ctx.ExecuteQuery()
                                                
                                                # If there are sharing links, try to remove them
                                                if ($sharingInfo.SharingLinks -and $sharingInfo.SharingLinks.Count -gt 0) {
                                                    foreach ($link in $sharingInfo.SharingLinks) {
                                                        if ($link.ShareId -and $groupName -like "*$($link.ShareId)*") {
                                                            # Found the matching link, try to delete it
                                                            $deleteResult = [Microsoft.SharePoint.Client.Sharing.WebSharingManager]::DeleteSharingLinkByUrl($ctx, $link.Url)
                                                            $ctx.Load($deleteResult)
                                                            $ctx.ExecuteQuery()
                                                            
                                                            if ($deleteResult.Value) {
                                                                Write-Host "        Successfully removed sharing link using CSOM WebSharingManager" -ForegroundColor Green
                                                                Write-InfoLog -LogName $Log -LogEntryText "Successfully removed sharing link using CSOM WebSharingManager"
                                                                $linkRemoved = $true
                                                                break
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch {
                                                Write-LogEntry -LogName $Log -LogEntryText "WebSharingManager approach failed: $_" -Level "DEBUG"
                                                
                                                # Final fallback: Try to break role inheritance and clear sharing
                                                try {
                                                    Write-DebugLog -LogName $Log -LogEntryText "Attempting to break inheritance and clear sharing"
                                                    
                                                    # Break role inheritance to stop sharing
                                                    $listItem.BreakRoleInheritance($false, $false)
                                                    $ctx.ExecuteQuery()
                                                    
                                                    Write-Host "        Successfully broke inheritance to stop sharing" -ForegroundColor Green
                                                    Write-LogEntry -LogName $Log -LogEntryText "Successfully broke inheritance to stop sharing" -Level "INFO"
                                                    $linkRemoved = $true
                                                }
                                                catch {
                                                    Write-LogEntry -LogName $Log -LogEntryText "Break inheritance also failed: $_" -Level "DEBUG"
                                                }
                                            }
                                        }
                                        catch {
                                            Write-LogEntry -LogName $Log -LogEntryText "CSOM approach completely failed: $_" -Level "DEBUG"
                                        }
                                    }
                                }
                                
                                # Return the result
                                return $linkRemoved
                            } -Operation "Unshare link using PnP and CSOM methods"
                            
                            # Set the result from the script block
                            $sharingLinkRemoved = $result
                            
                            if (-not $sharingLinkRemoved) {
                                Write-ErrorLog -LogName $Log -LogEntryText "All for group: $groupName"
                            }
                        }
                        catch {
                            Write-ErrorLog -LogName $Log -LogEntryText "All: $_"
                        }
                    }
                    
                    # Only try group removal if sharing link removal failed
                    if (-not $sharingLinkRemoved) {
                        Write-LogEntry -LogName $Log -LogEntryText "Sharing link removal failed, falling back to group removal method" -Level "INFO"
                        
                        try {
                            Invoke-WithThrottleHandling -ScriptBlock {
                                # Try Force parameter first, fallback to no confirmation parameter
                                try {
                                    # First check if group still exists
                                    $groupCheck = Get-PnPGroup -Identity $orgGroup.Id -ErrorAction SilentlyContinue
                                    if ($groupCheck) {
                                        Remove-PnPGroup -Identity $orgGroup.Id -Force
                                        Write-Host "      Successfully removed empty sharing group: $groupName" -ForegroundColor Green
                                        Write-InfoLog -LogName $Log -LogEntryText "Successfully removed empty Organization sharing group: $groupName"
                                    }
                                    else {
                                        Write-LogEntry -LogName $Log -LogEntryText "Group $groupName no longer exists, may have already been removed" -Level "INFO"
                                        Write-Host "      Group no longer exists (may have already been removed)" -ForegroundColor Yellow
                                    }
                                }
                                catch {
                                    # Fallback if Force parameter is not supported
                                    $groupCheck = Get-PnPGroup -Identity $orgGroup.Id -ErrorAction SilentlyContinue
                                    if ($groupCheck) {
                                        Remove-PnPGroup -Identity $orgGroup.Id
                                        Write-Host "      Successfully removed empty sharing group: $groupName" -ForegroundColor Green
                                        Write-InfoLog -LogName $Log -LogEntryText "Successfully removed empty Organization sharing group: $groupName"
                                    }
                                    else {
                                        Write-LogEntry -LogName $Log -LogEntryText "Group $groupName no longer exists during fallback removal" -Level "INFO"
                                        Write-Host "      Group no longer exists (fallback check)" -ForegroundColor Yellow
                                    }
                                }
                            } -Operation "Remove empty Organization sharing group $groupName"
                        }
                        catch {
                            $removeError = $_.Exception.Message
                            Write-LogEntry -LogName $Log -LogEntryText "Could not remove sharing group $groupName : $removeError" -Level "DEBUG"
                            
                            # If standard removal fails, try CSOM force removal only if error is not about missing group
                            if ($removeError -notmatch "does not exist|not found|Group cannot be found") {
                                try {
                                    Write-DebugLog -LogName $Log -LogEntryText "Attempting CSOM force removal of group $groupName"
                                    Invoke-WithThrottleHandling -ScriptBlock {
                                        # Get the web and group collection
                                        $web = Get-PnPWeb
                                        $ctx = $web.Context
                                        
                                        # First check if group exists in CSOM context
                                        try {
                                            $groupToRemove = $web.SiteGroups.GetById($orgGroup.Id)
                                            $ctx.Load($groupToRemove)
                                            $ctx.ExecuteQuery()
                                            
                                            # If we get here, group exists, so remove it
                                            $web.SiteGroups.Remove($groupToRemove)
                                            $ctx.ExecuteQuery()
                                            
                                            Write-Host "      Force removed sharing group: $groupName" -ForegroundColor Green
                                            Write-LogEntry -LogName $Log -LogEntryText "Force removed Organization sharing group: $groupName" -Level "INFO"
                                        }
                                        catch {
                                            if ($_.Exception.Message -like "*Group cannot be found*") {
                                                Write-LogEntry -LogName $Log -LogEntryText "Group $groupName no longer exists in CSOM context" -Level "INFO"
                                                Write-Host "      Group no longer exists (CSOM check)" -ForegroundColor Yellow
                                            }
                                            else {
                                                throw
                                            }
                                        }
                                    } -Operation "Force remove Organization sharing group using CSOM"
                                }
                                catch {
                                    Write-Host "      Warning: Could not remove sharing group $groupName : $_" -ForegroundColor Red
                                    Write-LogEntry -LogName $Log -LogEntryText "Final attempt failed to remove sharing group $groupName : $_" -Level "ERROR"
                                }
                            }
                            else {
                                Write-LogEntry -LogName $Log -LogEntryText "Group $groupName appears to have already been removed: $removeError" -Level "INFO"
                                Write-Host "      Group appears to have already been removed" -ForegroundColor Yellow
                            }
                        }
                    }
                    else {
                        # Sharing link was successfully removed, no need to remove group
                        Write-Host "      Sharing link successfully removed, group should be automatically cleaned up" -ForegroundColor Green
                        Write-LogEntry -LogName $Log -LogEntryText "Sharing link successfully removed for group $groupName, skipping manual group removal" -Level "INFO"
                    }
                    
                    # Update the link removal status in site collection data
                    Update-LinkRemovalStatus -SiteUrl $SiteUrl -SharingGroupName $groupName -WasRemoved $sharingLinkRemoved
                }
            }
            catch {
                Write-Host "      Warning: Error during sharing link removal for $groupName : $_" -ForegroundColor Red
                Write-ErrorLog -LogName $Log -LogEntryText "Error during $_"
                
                # Update status as failed for this group
                Update-LinkRemovalStatus -SiteUrl $SiteUrl -SharingGroupName $groupName -WasRemoved $false
            }
        }
    }
    catch {
        Write-Host "  Error processing Organization sharing links for site $SiteUrl : $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error during $_"
    }
}

# ----------------------------------------------
# Function to clean up corrupted sharing groups
# ----------------------------------------------
Function Remove-CorruptedSharingGroups {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl
    )
    
    Write-Host "  Checking for corrupted sharing groups on site: $SiteUrl" -ForegroundColor Yellow
    Write-LogEntry -LogName $Log -LogEntryText "Checking for corrupted sharing groups on site: $SiteUrl" -Level "INFO"
    
    try {
        # Connect to the specific site
        Connect-PnPOnline -Url $SiteUrl @connectionParams -ErrorAction Stop
        
        # Get all SharePoint groups that look like sharing groups
        $allSharingGroups = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPGroup | Where-Object { $_.Title -like "SharingLinks*" }
        } -Operation "Get all sharing groups for $SiteUrl"
        
        if ($allSharingGroups.Count -eq 0) {
            Write-LogEntry -LogName $Log -LogEntryText "No sharing groups found on site: $SiteUrl" -Level "DEBUG"
            return
        }
        
        $corruptedGroupsRemoved = 0
        
        foreach ($sharingGroup in $allSharingGroups) {
            try {
                # Check if group has any members
                $groupMembers = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPGroupMember -Identity $sharingGroup.Id -ErrorAction SilentlyContinue
                } -Operation "Check members in sharing group $($sharingGroup.Title)"
                
                # If group has no members, it's likely corrupted
                if (-not $groupMembers -or $groupMembers.Count -eq 0) {
                    Write-Host "    Found empty sharing group: $($sharingGroup.Title)" -ForegroundColor Yellow
                    Write-LogEntry -LogName $Log -LogEntryText "Found potentially corrupted empty sharing group: $($sharingGroup.Title)" -Level "INFO"
                    
                    try {
                        # Try to remove the empty sharing group
                        Invoke-WithThrottleHandling -ScriptBlock {
                            Remove-PnPGroup -Identity $sharingGroup.Id -Force
                        } -Operation "Remove corrupted sharing group $($sharingGroup.Title)"
                        
                        Write-Host "    Successfully removed corrupted sharing group: $($sharingGroup.Title)" -ForegroundColor Green
                        Write-InfoLog -LogName $Log -LogEntryText "Successfully removed corrupted sharing group: $($sharingGroup.Title)"
                        $corruptedGroupsRemoved++
                    }
                    catch {
                        Write-Host "    Warning: Could not remove corrupted sharing group: $($sharingGroup.Title) - $_" -ForegroundColor Red
                        Write-ErrorLog -LogName $Log -LogEntryText "Could not remove corrupted sharing group $($sharingGroup.Title): $_"
                    }
                }
            }
            catch {
                Write-ErrorLog -LogName $Log -LogEntryText "Error processing sharing group $($sharingGroup.Title): $_"
            }
        }
        
        if ($corruptedGroupsRemoved -gt 0) {
            Write-Host "  Removed $corruptedGroupsRemoved corrupted sharing groups from site: $SiteUrl" -ForegroundColor Green
            Write-InfoLog -LogName $Log -LogEntryText "Removed $corruptedGroupsRemoved corrupted sharing groups from site: $SiteUrl"
        }
    }
    catch {
        Write-Host "  Error processing corrupted sharing groups for site $SiteUrl : $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error processing corrupted sharing groups for site $SiteUrl : $_"
    }
}

# ----------------------------------------------
# Function to detect and parse script's CSV output for Organization links
# ----------------------------------------------
Function Test-AndParseScriptCsvOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string] $FilePath
    )
    
    try {
        # Read the first few lines to check the header format
        $firstLine = Get-Content -Path $FilePath -TotalCount 1
        
        # Check if this looks like our script's CSV output format
        $expectedHeaders = @("Site URL", "Site Owner", "IB Mode", "IB Segment", "Site Template", "Sharing Group Name", "Sharing Link Members", "File URL", "File Owner", "IsTeamsConnected", "SharingCapability", "Last Content Modified", "Link Removed")
        
        if ($firstLine -and $firstLine.Contains("Sharing Group Name")) {
            Write-Host "Detected script's CSV output format - will process Organization sharing links only" -ForegroundColor Cyan
            Write-InfoLog -LogName $Log -LogEntryText "Input file detected as script's CSV output format"
            
            # Import the full CSV
            $csvData = Import-Csv -Path $FilePath
            
            # Filter for Organization sharing links only
            $organizationEntries = $csvData | Where-Object { 
                $_."Sharing Group Name" -like "*Organization*" -and
                -not [string]::IsNullOrWhiteSpace($_."Site URL")
            }
            
            if ($organizationEntries.Count -eq 0) {
                Write-Host "No Organization sharing links found in the input CSV file" -ForegroundColor Yellow
                Write-InfoLog -LogName $Log -LogEntryText "No Organization sharing links found in input CSV"
                return @{
                    IsScriptOutput    = $true
                    Sites             = @()
                    OrganizationLinks = @{
                    }
                }
            }
            
            # Group by Site URL to get unique sites
            $siteGroups = $organizationEntries | Group-Object "Site URL"
            
            # Create sites collection for processing
            $sitesToProcess = @()
            $orgLinksData = @{
            }
            
            foreach ($siteGroup in $siteGroups) {
                $siteUrl = $siteGroup.Name
                $sitesToProcess += [PSCustomObject]@{ URL = $siteUrl }
                
                # Store Organization sharing group details for this site
                $orgLinksData[$siteUrl] = @{
                    Groups               = @()
                    HasOrganizationLinks = $true
                }
                
                foreach ($entry in $siteGroup.Group) {
                    $orgLinksData[$siteUrl].Groups += @{
                        GroupName = $entry."Sharing Group Name"
                        Members   = $entry."Sharing Link Members"
                        FileUrl   = $entry."File URL"
                        FileOwner = $entry."File Owner"
                    }
                }
            }
            
            Write-Host "Found $($sitesToProcess.Count) sites with Organization sharing links for remediation" -ForegroundColor Green
            Write-InfoLog -LogName $Log -LogEntryText "Parsed $($sitesToProcess.Count) sites with Organization sharing links from CSV input"
            
            return @{
                IsScriptOutput    = $true
                Sites             = $sitesToProcess
                OrganizationLinks = $orgLinksData
            }
        }
        else {
            Write-Host "Input file appears to be a simple site URL list" -ForegroundColor Yellow
            Write-InfoLog -LogName $Log -LogEntryText "Input file detected as simple site URL list"
            
            return @{
                IsScriptOutput    = $false
                Sites             = @()
                OrganizationLinks = @{
                }
            }
        }
    }
    catch {
        Write-ErrorLog -LogName $Log -LogEntryText "Error analyzing input file format: $_"
        throw "Error analyzing input file format: $_"
    }
}

# Main Processing Loop
# ----------------------------------------------
$totalSites = $sites.Count
$processedCount = 0
$sitesWithSharingLinksCount = 0
$organizationLinksProcessedCount = 0

foreach ($site in $sites) {
    $processedCount++
    $siteUrl = ""
    
    # Handle both input file format and Get-PnPTenantSite format
    if ($site.URL) {
        $siteUrl = $site.URL
    }
    elseif ($site.Url) {
        $siteUrl = $site.Url
    }
    else {
        $siteUrl = $site.ToString()
    }
    
    # Skip if empty URL
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        continue
    }
    
    Write-Host "Processing site $processedCount of $totalSites : $siteUrl" -ForegroundColor Green
    Write-InfoLog -LogName $Log -LogEntryText "Processing site $processedCount of $totalSites : $siteUrl"
    
    try {
        # Connect to the specific site to get groups and users
        try {
            Connect-PnPOnline -Url $siteUrl @connectionParams -ErrorAction Stop
            
            # Get Site Properties using SharePoint Admin connection
            Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop
            $siteProperties = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPTenantSite -Identity $siteUrl
            } -Operation "Get site properties for $siteUrl"
            
            # Connect back to the site for group processing
            Connect-PnPOnline -Url $siteUrl @connectionParams -ErrorAction Stop
            
            # Initialize site data
            Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteProperties
            
            # Get all groups for this site
            $spGroups = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPGroup
            } -Operation "Get groups for site $siteUrl"
            
            foreach ($spGroup in $spGroups) {
                $spGroupName = $spGroup.Title
                
                # Update site data with group information
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteProperties -SPGroupName $spGroupName
                
                # Get users in each group
                $spUsers = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPGroupMember -Identity $spGroup.Id
                } -Operation "Get members for group $spGroupName"
                
                foreach ($spUser in $spUsers) {
                    if ($spUser -and $spUser.LoginName) {
                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteProperties -AssociatedSPGroup $spGroupName -SPUserName $spUser.LoginName -SPUserTitle $spUser.Title -SPUserEmail $spUser.Email
                    }
                }
                
                # Extract document information from sharing groups
                if ($spGroupName -like "SharingLinks*") {
                    try {
                        # Extract document ID from sharing group name
                        if ($spGroupName -match "SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.") {
                            $documentId = $matches[1]
                            $sharingType = "Unknown"
                            $documentUrl = ""
                            $documentOwner = ""
                            
                            # Determine sharing type from group name
                            if ($spGroupName -like "*OrganizationView*") {
                                $sharingType = "OrganizationView"
                            }
                            elseif ($spGroupName -like "*OrganizationEdit*") {
                                $sharingType = "OrganizationEdit"
                            }
                            elseif ($spGroupName -like "*AnonymousAccess*") {
                                $sharingType = "AnonymousAccess"
                            }
                            
                            # Try to find the document using Microsoft Graph
                            try {
                                $graphToken = Invoke-WithThrottleHandling -ScriptBlock {
                                    Get-PnPAccessToken
                                } -Operation "Get-PnPAccessToken for document search"
                                
                                if ($graphToken) {
                                    $headers = @{
                                        "Authorization" = "Bearer $graphToken"
                                        "Content-Type"  = "application/json"
                                    }
                                    
                                    $searchQuery = @{
                                        requests = @(
                                            @{
                                                entityTypes               = @("driveItem")
                                                query                     = @{
                                                    queryString = "UniqueID:$documentId"
                                                }
                                                from                      = 0
                                                size                      = 25
                                                sharePointOneDriveOptions = @{
                                                    includeContent = "sharedContent,privateContent"
                                                }
                                                region                    = $searchRegion
                                            }
                                        )
                                    }
                                    
                                    $searchBody = $searchQuery | ConvertTo-Json -Depth 5
                                    $searchUrl = "https://graph.microsoft.com/v1.0/search/query"
                                    
                                    $searchResults = Invoke-WithThrottleHandling -ScriptBlock {
                                        Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body $searchBody
                                    } -Operation "Microsoft Graph Search for document ID $documentId"
                                    
                                    if ($searchResults.value -and 
                                        $searchResults.value[0].hitsContainers -and 
                                        $searchResults.value[0].hitsContainers[0].hits -and 
                                        $searchResults.value[0].hitsContainers[0].hits.Count -gt 0) {
                                        
                                        $hit = $searchResults.value[0].hitsContainers[0].hits[0]
                                        $resource = $hit.resource
                                        
                                        if ($resource) {
                                            # Use the WebUrl from the result if available
                                            if ($resource.webUrl) {
                                                $documentUrl = $resource.webUrl
                                            }
                                            
                                            # Try to get the author/owner if available
                                            if ($resource.createdBy.user.displayName) {
                                                $ownerDisplayName = $resource.createdBy.user.displayName
                                                $ownerEmail = $resource.createdBy.user.email
                                                
                                                if ($ownerEmail) {
                                                    $documentOwner = "$ownerDisplayName <$ownerEmail>"
                                                }
                                                else {
                                                    $documentOwner = $ownerDisplayName
                                                }
                                            }
                                            
                                            Write-LogEntry -LogName $Log -LogEntryText "Located document via Graph search: $documentUrl" -Level "DEBUG"
                                        }
                                    }
                                    else {
                                        Write-LogEntry -LogName $Log -LogEntryText "No matching documents found in Graph search for ID: $documentId" -Level "DEBUG"
                                    }
                                }
                                else {
                                    Write-LogEntry -LogName $Log -LogEntryText "Unable to get Graph access token for document search." -Level "ERROR"
                                }
                            }
                            catch {
                                Write-ErrorLog -LogName $Log -LogEntryText "Error searching for document via Graph API: ${_}"
                            }
                            
                            # Store the sharing link information 
                            if (-not $siteCollectionData[$siteUrl].ContainsKey("DocumentDetails")) {
                                $siteCollectionData[$siteUrl]["DocumentDetails"] = @{
                                }
                            }
                            
                            $siteCollectionData[$siteUrl]["DocumentDetails"][$spGroupName] = @{
                                "DocumentId"    = $documentId
                                "SharingType"   = $sharingType
                                "DocumentUrl"   = $documentUrl
                                "DocumentOwner" = $documentOwner
                                "SharedOn"      = $siteUrl
                            }
                            
                            Write-DebugLog -LogName $Log -LogEntryText "Stored sharing information for document ID $documentId"
                        }
                    }
                    catch {
                        Write-ErrorLog -LogName $Log -LogEntryText "Error extracting document ID from group name $($spGroupName) : ${_}"
                    }
                }
            }

            # Process and write sharing links data for this site immediately if any found
            if ($siteCollectionData[$siteUrl]["Has Sharing Links"]) {
                $sitesWithSharingLinksCount++
                
                # Convert Organization sharing links to direct permissions if enabled
                if ($convertOrganizationLinks) {
                    Convert-OrganizationSharingLinks -SiteUrl $siteUrl
                    $organizationLinksProcessedCount++
                }
                
                # Clean up any remaining corrupted sharing groups if enabled
                if ($cleanupCorruptedSharingGroups) {
                    try {
                        Remove-CorruptedSharingGroups -SiteUrl $siteUrl
                    }
                    catch {
                        Write-ErrorLog -LogName $Log -LogEntryText "Error during final cleanup for site $siteUrl : $_"
                    }
                }
                
                # Write sharing links data AFTER processing Organization links and cleanup
                # This ensures the "Link Removed" status is accurate
                Write-SiteSharingLinks -SiteUrl $siteUrl -SiteData $siteCollectionData[$siteUrl]
            }
        }
        catch { 
            Write-ErrorLog -LogName $Log -LogEntryText "Could not connect to site $siteUrl : ${_}" 
        }
    }
    catch {
        Write-ErrorLog -LogName $Log -LogEntryText "Error processing site $siteUrl : ${_}"
        continue
    }
}

# ----------------------------------------------
# Final Output Generation
# ----------------------------------------------
Write-Host "Consolidating results..." -ForegroundColor Green

# No incremental file generation - only focus on sharing links output
if ($sitesWithSharingLinksCount -gt 0) {
    Write-Host "Found $sitesWithSharingLinksCount site collections with sharing links" -ForegroundColor Green
    Write-Host "Sharing links data written to: $sharingLinksOutputFile" -ForegroundColor Green
    Write-InfoLog -LogName $Log -LogEntryText "Total sites with sharing links: $sitesWithSharingLinksCount"
    
    if ($convertOrganizationLinks) {
        Write-Host "Processed Organization sharing links on $organizationLinksProcessedCount sites" -ForegroundColor Green
        Write-InfoLog -LogName $Log -LogEntryText "Processed Organization sharing links on $organizationLinksProcessedCount sites"
    }
}
else {
    Write-Host "No site collections with sharing links found." -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "No site collections with sharing links found."
}

# ----------------------------------------------
# Disconnect and finish
# ----------------------------------------------
Disconnect-PnPOnline
Write-InfoLog -LogName $Log -LogEntryText "Script finished."
Write-Host "Script finished. Log file located at: $log" -ForegroundColor Green
