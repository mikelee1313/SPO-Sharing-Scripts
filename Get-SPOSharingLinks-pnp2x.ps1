<#
.SYNOPSIS
    Detects and inventories SharePoint Online sharing links across the tenant (Detection Mode Only).

.DESCRIPTION
    This lightweight script scans SharePoint Online sites to identify and inventory all sharing links, with a focus on Organization sharing links. 
    This script is DETECTION ONLY and will never modify any permissions, remove sharing links, or clean up sharing groups.
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

.PARAMETER debugLogging
    When set to $true, the script logs detailed DEBUG operations for troubleshooting.
    When set to $false, only INFO and ERROR operations are logged.

.PARAMETER inputfile
    Optional. Path to a CSV file containing SharePoint site URLs (one URL per line or with "URL" header).
    If not specified, the script will process all sites in the tenant.

.OUTPUTS
    - CSV file containing detailed information about sharing links found
    - Log file with operation details and errors

.NOTES
    Authors: Mike Lee
    Updated: 8/7/2025

    - Requires PnP.PowerShell 2.x module
    - Requires an Entra app registration with appropriate SharePoint permissions
       - The app must have:
        - Sharepoint:Sites.FullControl.All
        - SharePoint:User.Read.All 
        - Graph:Sites.FullControl.All
        - Graph:Sites.Read.All
        - Graph:Files.Read.All
    - Requires a certificate for authentication
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
    # Inventory all sharing links from all sites in the tenant
    .\Get-SPOSharingLinks-pnp2x.ps1

.EXAMPLE
    # Inventory sharing links from specific sites listed in a CSV file
    $inputfile = "C:\temp\MySites.csv"
    .\Get-SPOSharingLinks-pnp2x.ps1
#>

# ----------------------------------------------
# Set Variables
# ----------------------------------------------
$tenantname = "m365x61250205"                                   # This is your tenant name
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51"                # This is your Tenant ID
$searchRegion = "NAM"                                           # Region for Microsoft Graph search
$debugLogging = $false                                          # Set to $true for detailed DEBUG logging, $false for INFO and ERROR logging only

# ----------------------------------------------
# Initialize Parameters - Do not change
# ----------------------------------------------
$sites = @()
$inputfile = $null
$log = $null
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# ----------------------------------------------
# Input / Output and Log Files
# ----------------------------------------------
$inputfile = "C:\temp\oversharedurls - Copy.txt" #If no input file specified, will process all sites in the tenant
$log = "$env:TEMP\" + 'SPOSharingLinks_DetectionOnly_' + $date + '_' + "logfile.log"
# Initialize sharing links output file
$sharingLinksOutputFile = "$env:TEMP\" + 'SPO_SharingLinks_DetectionOnly_' + $date + '.csv'

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
        if ($null -ne $LogName) {
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
# Script Mode Information
# ----------------------------------------------
Write-Host "Script is running in DETECTION-ONLY mode" -ForegroundColor Cyan
Write-InfoLog -LogName $Log -LogEntryText "Script is running in DETECTION-ONLY mode - Only detecting and inventorying sharing links, no modifications will be made"

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
        # Simple site URL list handling
        Write-Host "Input file appears to be a simple site URL list" -ForegroundColor Yellow
        Write-InfoLog -LogName $Log -LogEntryText "Input file detected as simple site URL list"
        $sites = Import-csv -path $inputfile -Header 'URL'
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
$sharingLinksHeaders = "Site URL,Site Owner,IB Mode,IB Segment,Site Template,Sharing Group Name,Sharing Link Members,File URL,File Owner,Filename,SharingType,Sharing Link URL,Link Expiration Date,IsTeamsConnected,SharingCapability,Last Content Modified,Link Removed"
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
        }
    }

    # Check for SharingLinks groups
    if (-not [string]::IsNullOrWhiteSpace($SPGroupName) -and $SPGroupName -like "SharingLinks*") {
        $siteCollectionData[$SiteUrl]["Has Sharing Links"] = $true
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
                $documentItemType = "Not found"
                $sharingLinkUrl = "Not found"
                $linkExpirationDate = "Not found"
                if ($SiteData.ContainsKey("DocumentDetails") -and $SiteData["DocumentDetails"].ContainsKey($sharingGroup)) {
                    $documentUrl = $SiteData["DocumentDetails"][$sharingGroup]["DocumentUrl"]
                    $documentOwner = $SiteData["DocumentDetails"][$sharingGroup]["DocumentOwner"]
                    $documentItemType = $SiteData["DocumentDetails"][$sharingGroup]["DocumentItemType"]
                    $sharingLinkUrl = $SiteData["DocumentDetails"][$sharingGroup]["SharingLinkUrl"]
                    $linkExpirationDate = $SiteData["DocumentDetails"][$sharingGroup]["ExpirationDate"]
                    Write-DebugLog -LogName $Log -LogEntryText "Retrieved document details for $sharingGroup - URL: $documentUrl, Owner: $documentOwner, Type: $documentItemType, LinkURL: $sharingLinkUrl, Expiration: $linkExpirationDate"
                }
                else {
                    Write-DebugLog -LogName $Log -LogEntryText "No document details found for sharing group: $sharingGroup. DocumentDetails exists: $($SiteData.ContainsKey('DocumentDetails')), Group key exists: $(if ($SiteData.ContainsKey('DocumentDetails')) { $SiteData['DocumentDetails'].ContainsKey($sharingGroup) } else { 'N/A' })"
                }
                
                # For detection-only mode, always set Link Removed to "False" since we never remove anything
                $linkRemoved = "False"
                
                # Extract filename from the document URL
                $filename = "Not found"
                if ($documentUrl -ne "Not found" -and -not [string]::IsNullOrWhiteSpace($documentUrl)) {
                    try {
                        if ($documentUrl -match "DispForm\.aspx\?ID=(\d+)") {
                            # This is a list item - try to get a meaningful name
                            # For list items, we'll use "List Item" + ID as the filename
                            $itemId = $matches[1]
                            $filename = "List Item $itemId"
                            
                            # Try to extract list name for better context
                            if ($documentUrl -match "/Lists/([^/]+)/DispForm\.aspx") {
                                $listName = $matches[1]
                                $filename = "$listName - Item $itemId"
                            }
                        }
                        else {
                            # This is a regular file - extract filename from URL
                            $uri = [System.Uri]$documentUrl
                            $pathParts = $uri.AbsolutePath.Split('/')
                            $filename = $pathParts[$pathParts.Length - 1]
                            
                            # Decode URL encoding if present
                            $filename = [System.Web.HttpUtility]::UrlDecode($filename)
                        }
                    }
                    catch {
                        Write-DebugLog -LogName $Log -LogEntryText "Could not extract filename from URL: $documentUrl. Error: $_"
                        $filename = "Extraction Error"
                    }
                }
                
                # Determine sharing type based on sharing group name
                $sharingType = "Unknown"
                if ($sharingGroup -like "*Flexible*") {
                    $sharingType = "Flexible"
                }
                elseif ($sharingGroup -like "*Organization*") {
                    $sharingType = "Organization"
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
                    "Filename"              = $filename
                    "SharingType"           = $sharingType
                    "Sharing Link URL"      = $sharingLinkUrl
                    "Link Expiration Date"  = $linkExpirationDate
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
# Function to search for documents using Microsoft Graph API
# ----------------------------------------------
Function Search-DocumentViaGraphAPI {
    param(
        [Parameter(Mandatory = $true)]
        [string] $DocumentId,
        [Parameter(Mandatory = $true)]
        [string] $SearchRegion,
        [Parameter(Mandatory = $false)]
        [string] $LogContext = "Document search"
    )
    
    $result = @{
        Found         = $false
        DocumentUrl   = ""
        DocumentOwner = ""
        ItemType      = ""
    }
    
    try {
        Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Searching for document ID: $DocumentId"
        
        $graphToken = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPGraphAccessToken
        } -Operation "Get-PnPGraphAccessToken for $LogContext"
        
        if (-not $graphToken) {
            Write-ErrorLog -LogName $Log -LogEntryText "$LogContext - Unable to get Graph access token"
            return $result
        }
        
        $headers = @{
            "Authorization" = "Bearer $graphToken"
            "Content-Type"  = "application/json"
        }
        
        $searchUrl = "https://graph.microsoft.com/v1.0/search/query"
        $itemFound = $false
        
        # First, try searching as driveItem (files in document libraries)
        $driveItemSearchQuery = @{
            requests = @(
                @{
                    entityTypes               = @("driveItem")
                    query                     = @{
                        queryString = "UniqueID:$DocumentId"
                    }
                    from                      = 0
                    size                      = 25
                    sharePointOneDriveOptions = @{
                        includeContent = "sharedContent,privateContent"
                    }
                    region                    = $SearchRegion
                }
            )
        }
        
        $searchBody = $driveItemSearchQuery | ConvertTo-Json -Depth 5
        
        $searchResults = Invoke-WithThrottleHandling -ScriptBlock {
            Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body $searchBody
        } -Operation "$LogContext - Microsoft Graph Search for document ID $DocumentId (driveItem)"
        
        Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Graph search response structure: value count = $(if ($searchResults.value) { $searchResults.value.Count } else { 'null' })"
        
        # Check if we found results with driveItem search
        if ($searchResults.value -and 
            $searchResults.value[0].hitsContainers -and 
            $searchResults.value[0].hitsContainers[0].hits -and 
            $searchResults.value[0].hitsContainers[0].hits.Count -gt 0) {
            
            $hit = $searchResults.value[0].hitsContainers[0].hits[0]
            $resource = $hit.resource
            
            Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Found hit with resource. WebUrl: '$($resource.webUrl)', CreatedBy: '$($resource.createdBy.user.displayName)'"
            
            if ($resource) {
                $result.Found = $true
                $result.ItemType = "driveItem"
                
                # Use the WebUrl from the result if available
                if ($resource.webUrl) {
                    $result.DocumentUrl = $resource.webUrl
                }
                
                # Try to get the author/owner if available
                if ($resource.createdBy.user.displayName) {
                    $ownerDisplayName = $resource.createdBy.user.displayName
                    $ownerEmail = $resource.createdBy.user.email
                    
                    if ($ownerEmail) {
                        $result.DocumentOwner = "$ownerDisplayName <$ownerEmail>"
                    }
                    else {
                        $result.DocumentOwner = $ownerDisplayName
                    }
                }
                
                Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Located document via Graph search (driveItem): $($result.DocumentUrl)"
                $itemFound = $true
            }
        }
        
        # If no results found with driveItem, try searching as listItem (SharePoint list items)
        if (-not $itemFound) {
            Write-DebugLog -LogName $Log -LogEntryText "$LogContext - No driveItem found, trying listItem search for ID: $DocumentId"
            
            $listItemSearchQuery = @{
                requests = @(
                    @{
                        entityTypes               = @("listItem")
                        query                     = @{
                            queryString = "UniqueID:$DocumentId"
                        }
                        from                      = 0
                        size                      = 25
                        sharePointOneDriveOptions = @{
                            includeContent = "sharedContent,privateContent"
                        }
                        region                    = $SearchRegion
                    }
                )
            }
            
            $listItemSearchBody = $listItemSearchQuery | ConvertTo-Json -Depth 5
            
            $listItemSearchResults = Invoke-WithThrottleHandling -ScriptBlock {
                Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body $listItemSearchBody
            } -Operation "$LogContext - Microsoft Graph Search for document ID $DocumentId (listItem)"
            
            Write-DebugLog -LogName $Log -LogEntryText "$LogContext - List item search response structure: value count = $(if ($listItemSearchResults.value) { $listItemSearchResults.value.Count } else { 'null' })"
            
            if ($listItemSearchResults.value -and 
                $listItemSearchResults.value[0].hitsContainers -and 
                $listItemSearchResults.value[0].hitsContainers[0].hits -and 
                $listItemSearchResults.value[0].hitsContainers[0].hits.Count -gt 0) {
                
                $hit = $listItemSearchResults.value[0].hitsContainers[0].hits[0]
                $resource = $hit.resource
                
                Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Found list item hit with resource. WebUrl: '$($resource.webUrl)', CreatedBy: '$($resource.createdBy.user.displayName)'"
                
                if ($resource) {
                    $result.Found = $true
                    $result.ItemType = "listItem"
                    
                    # Use the WebUrl from the result if available
                    if ($resource.webUrl) {
                        $result.DocumentUrl = $resource.webUrl
                    }
                    
                    # Try to get the author/owner if available
                    if ($resource.createdBy.user.displayName) {
                        $ownerDisplayName = $resource.createdBy.user.displayName
                        $ownerEmail = $resource.createdBy.user.email
                        
                        if ($ownerEmail) {
                            $result.DocumentOwner = "$ownerDisplayName <$ownerEmail>"
                        }
                        else {
                            $result.DocumentOwner = $ownerDisplayName
                        }
                    }
                    
                    Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Located list item via Graph search (listItem): $($result.DocumentUrl)"
                    $itemFound = $true
                }
            }
            else {
                Write-DebugLog -LogName $Log -LogEntryText "$LogContext - No matching items found in Graph search (both driveItem and listItem) for ID: $DocumentId"
            }
        }
    }
    catch {
        Write-ErrorLog -LogName $Log -LogEntryText "$LogContext - Error searching for document via Graph API: $_"
    }
    
    return $result
}

# ----------------------------------------------
# Function to collect and store sharing link URLs for a site
# ----------------------------------------------
Function Get-SharingLinkUrls {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl
    )
    
    Write-Host "  Collecting sharing link URLs for site: $SiteUrl" -ForegroundColor Cyan
    Write-InfoLog -LogName $Log -LogEntryText "Collecting sharing link URLs for site: $SiteUrl"
    
    try {
        # Connect to the specific site if not already connected
        $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if (-not $currentConnection -or $currentConnection.Url -ne $SiteUrl) {
            Connect-PnPOnline -Url $SiteUrl @connectionParams -ErrorAction Stop
        }
        
        # Get all SharingLinks groups
        $sharingGroups = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPGroup | Where-Object { $_.Title -like "SharingLinks*" }
        } -Operation "Get sharing groups for $SiteUrl"
        
        if ($sharingGroups.Count -eq 0) {
            Write-DebugLog -LogName $Log -LogEntryText "No sharing groups found on site: $SiteUrl"
            return
        }
        
        Write-Host "    Found $($sharingGroups.Count) sharing groups" -ForegroundColor Green
        Write-InfoLog -LogName $Log -LogEntryText "Found $($sharingGroups.Count) sharing groups on site: $SiteUrl"
        
        foreach ($group in $sharingGroups) {
            $groupName = $group.Title
            
            # Extract document ID from group name
            $documentId = ""
            if ($groupName -match "SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.") {
                $documentId = $matches[1]
                Write-DebugLog -LogName $Log -LogEntryText "Processing sharing group: $groupName with document ID: $documentId"
                
                # Check if we have document details for this group
                if ($siteCollectionData[$SiteUrl].ContainsKey("DocumentDetails") -and 
                    $siteCollectionData[$SiteUrl]["DocumentDetails"].ContainsKey($groupName) -and
                    -not [string]::IsNullOrWhiteSpace($siteCollectionData[$SiteUrl]["DocumentDetails"][$groupName]["DocumentUrl"])) {
                    
                    $docUrl = $siteCollectionData[$SiteUrl]["DocumentDetails"][$groupName]["DocumentUrl"]
                    
                    # Try to get sharing links using the document ID directly (works for both files and list items)
                    try {
                        $sharingLinks = $null
                        $sharingLinkUrl = "Not found"
                        $expirationDate = "No expiration"
                        
                        Write-DebugLog -LogName $Log -LogEntryText "Attempting to get sharing links using document ID: $documentId"
                        
                        # Use Get-PnPFileSharingLink with the document ID directly (works for both files and list items)
                        $sharingLinks = Invoke-WithThrottleHandling -ScriptBlock {
                            Get-PnPFileSharingLink -Identity $documentId -ErrorAction SilentlyContinue
                        } -Operation "Get sharing links for document ID: $documentId"
                        
                        if ($sharingLinks -and $sharingLinks.Count -gt 0) {
                            Write-DebugLog -LogName $Log -LogEntryText "Found $($sharingLinks.Count) sharing links for document"
                            
                            # Log the structure of the first sharing link object to help with debugging
                            if ($debugLogging -and $sharingLinks[0]) {
                                $firstLink = $sharingLinks[0]
                                Write-DebugLog -LogName $Log -LogEntryText "Sharing link object properties: $(($firstLink | Get-Member -MemberType Property).Name -join ', ')"
                                
                                if ($firstLink.link) {
                                    Write-DebugLog -LogName $Log -LogEntryText "Link property exists. Link properties: $(($firstLink.link | Get-Member -MemberType Property).Name -join ', ')"
                                    if ($firstLink.link.WebUrl) {
                                        Write-DebugLog -LogName $Log -LogEntryText "WebUrl found: $($firstLink.link.WebUrl)"
                                    }
                                    if ($firstLink.link.ExpirationDateTime) {
                                        Write-DebugLog -LogName $Log -LogEntryText "ExpirationDateTime found: $($firstLink.link.ExpirationDateTime)"
                                    }
                                }
                                else {
                                    Write-DebugLog -LogName $Log -LogEntryText "Link property doesn't exist or is null"
                                }
                                
                                # Also check for expiration date at the top level
                                if ($firstLink.ExpirationDateTime) {
                                    Write-DebugLog -LogName $Log -LogEntryText "Top-level ExpirationDateTime found: $($firstLink.ExpirationDateTime)"
                                }
                            }
                            
                            # Look for a matching sharing link
                            $matchingLink = $sharingLinks | Where-Object { $_.Id -and $groupName -like "*$($_.Id)*" } | Select-Object -First 1
                            
                            if ($matchingLink) {
                                # Get the WebUrl property of the sharing link from the link property
                                $sharingLinkUrl = if ($matchingLink.link -and $matchingLink.link.WebUrl) { 
                                    $matchingLink.link.WebUrl 
                                }
                                else { 
                                    "Not found" 
                                }
                                
                                # Get the expiration date of the sharing link
                                if ($matchingLink.link -and $matchingLink.link.ExpirationDateTime) {
                                    # Format the expiration date to a readable format
                                    try {
                                        $expDate = [DateTime]::Parse($matchingLink.link.ExpirationDateTime)
                                        $expirationDate = $expDate.ToString("yyyy-MM-dd HH:mm:ss")
                                    }
                                    catch {
                                        $expirationDate = $matchingLink.link.ExpirationDateTime
                                        Write-DebugLog -LogName $Log -LogEntryText "Could not parse expiration date: $($matchingLink.link.ExpirationDateTime)"
                                    }
                                }
                                elseif ($matchingLink.ExpirationDateTime) {
                                    # Alternative location for expiration date
                                    try {
                                        $expDate = [DateTime]::Parse($matchingLink.ExpirationDateTime)
                                        $expirationDate = $expDate.ToString("yyyy-MM-dd HH:mm:ss")
                                    }
                                    catch {
                                        $expirationDate = $matchingLink.ExpirationDateTime
                                        Write-DebugLog -LogName $Log -LogEntryText "Could not parse expiration date: $($matchingLink.ExpirationDateTime)"
                                    }
                                }
                                
                                # Store the results
                                $siteCollectionData[$SiteUrl]["DocumentDetails"][$groupName]["SharingLinkUrl"] = $sharingLinkUrl
                                $siteCollectionData[$SiteUrl]["DocumentDetails"][$groupName]["ExpirationDate"] = $expirationDate
                                
                                Write-DebugLog -LogName $Log -LogEntryText "Found sharing link URL for group $groupName - URL: $sharingLinkUrl, Expiration: $expirationDate"
                            }
                            else {
                                Write-DebugLog -LogName $Log -LogEntryText "No matching sharing link found for group: $groupName"
                            }
                        }
                        else {
                            Write-DebugLog -LogName $Log -LogEntryText "No sharing links found for document at: $docUrl"
                        }
                    }
                    catch {
                        Write-DebugLog -LogName $Log -LogEntryText "Error getting sharing links for document: $_"
                    }
                }
                else {
                    Write-DebugLog -LogName $Log -LogEntryText "No document details found for group: $groupName"
                }
            }
        }
    }
    catch {
        Write-Host "  Error collecting sharing link URLs for site $SiteUrl : $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error collecting sharing link URLs for site $SiteUrl : $_"
    }
}

# Main Processing Loop
# ----------------------------------------------
$totalSites = $sites.Count
$processedCount = 0
$sitesWithSharingLinksCount = 0

# Display script mode information before starting site processing
Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "SCRIPT MODE: DETECTION-ONLY" -ForegroundColor Cyan
Write-Host "  - Only DETECTING and INVENTORYING sharing links" -ForegroundColor Cyan
Write-Host "  - NO modifications will be made to permissions or sharing links" -ForegroundColor Cyan
Write-Host "  - Results will be saved to: $sharingLinksOutputFile" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""
Write-InfoLog -LogName $Log -LogEntryText "Starting to process $totalSites sites in DETECTION-ONLY mode"

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
                            Write-DebugLog -LogName $Log -LogEntryText "Extracted document ID: $documentId from sharing group: $spGroupName"
                            $sharingType = "Unknown"
                            $documentUrl = ""
                            $documentOwner = ""
                            $documentItemType = ""
                            
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
                                    Get-PnPGraphAccessToken
                                } -Operation "Get-PnPGraphAccessToken for document search"
                                
                                if ($graphToken) {
                                    $headers = @{
                                        "Authorization" = "Bearer $graphToken"
                                        "Content-Type"  = "application/json"
                                    }
                                    
                                    # Try to find the document via Microsoft Graph search using the document ID
                                    $searchResult = Search-DocumentViaGraphAPI -DocumentId $documentId -SearchRegion $searchRegion -LogContext "Main processing loop - document search"
                                    
                                    Write-DebugLog -LogName $Log -LogEntryText "Search result for document ID $documentId - Found: $($searchResult.Found), URL: '$($searchResult.DocumentUrl)', Owner: '$($searchResult.DocumentOwner)', Type: '$($searchResult.ItemType)'"
                                    
                                    if ($searchResult.Found) {
                                        if ($searchResult.DocumentUrl) {
                                            $documentUrl = $searchResult.DocumentUrl
                                        }
                                        
                                        if ($searchResult.DocumentOwner) {
                                            $documentOwner = $searchResult.DocumentOwner
                                        }
                                        
                                        if ($searchResult.ItemType) {
                                            $documentItemType = $searchResult.ItemType
                                        }
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
                                "DocumentId"       = $documentId
                                "SharingType"      = $sharingType
                                "DocumentUrl"      = $documentUrl
                                "DocumentOwner"    = $documentOwner
                                "DocumentItemType" = $documentItemType
                                "SharedOn"         = $siteUrl
                                "SharingLinkUrl"   = "" # Will be populated when processing sharing links
                                "ExpirationDate"   = "" # Will be populated when processing sharing links
                            }
                            
                            Write-DebugLog -LogName $Log -LogEntryText "Stored sharing information for document ID $documentId with URL: $documentUrl and Owner: $documentOwner"
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
                
                # Collect sharing link URLs for this site
                Get-SharingLinkUrls -SiteUrl $siteUrl
                
                # Write sharing links data for this site
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

if ($sitesWithSharingLinksCount -gt 0) {
    Write-Host "Found $sitesWithSharingLinksCount site collections with sharing links" -ForegroundColor Green
    Write-Host "Sharing links data written to: $sharingLinksOutputFile" -ForegroundColor Green
    Write-InfoLog -LogName $Log -LogEntryText "Total sites with sharing links: $sitesWithSharingLinksCount"
    Write-Host "  Mode: DETECTION-ONLY - No modifications were made to permissions or sharing links" -ForegroundColor Cyan
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
