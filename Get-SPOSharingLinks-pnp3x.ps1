<#
.SYNOPSIS
    SharePoint Online Sharing Information Collection Tool

.DESCRIPTION
    This script collects comprehensive information about SharePoint Online sites and their sharing configurations,
    with a special focus on external sharing links. It connects to SharePoint Online using app-only authentication
    and processes sites either from a CSV file or by retrieving all sites in the tenant.

.PARAMETER None
    This script uses predefined variables that need to be set at the beginning of the script.

.NOTES
    File Name      : Get-SPOSharingLinks-pnp3x.ps1
    Author         : Mike Lee
    Date Created   : 5/27/25
    Prerequisite   : 
    -    PnP PowerShell module (Tested with PNP 3.1.0)
    -    Microsoft Graph API permissions for app-only authentication
            Graph: Files.Read.All (Application)
            SharePoint: Sites.FullControl.All (Application) 
    -    App registration in Entra ID with certificate

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
    .\Get-SPOSharingLinks.ps1
    
    Runs the script with the configured parameters to collect SharePoint site sharing information.

.INPUTS
    Optional CSV file with site URLs (sitelist.csv) in the format:
    https://tenant.sharepoint.com/sites/site1
    https://tenant.sharepoint.com/sites/site2

.OUTPUTS
    1. Main output CSV file - Contains basic information for all sites
    2. Sharing links specific CSV file - Contains detailed information about sharing links including
       document URLs, members with access, and document owners
    3. Log file - Records the script's execution progress and any errors

.FUNCTIONALITY
    - Connects to SharePoint Online using app-only authentication
    - Collects site information including Information Barrier settings
    - Identifies sites with sharing links
    - Extracts document IDs from sharing link groups
    - Uses Microsoft Graph to locate documents being shared
    - Identifies users with access to shared documents
    - Consolidates and exports data in structured CSV format

.NOTES
    Required variables to configure before running:
    - $tenantname : Your SharePoint tenant name (without .sharepoint.com)
    - $appID : Entra App ID for authentication
    - $thumbprint : Certificate thumbprint for app-only authentication
    - $tenant : Tenant ID (GUID)

    The script requires the PnP PowerShell module and appropriate permissions
    for app-only authentication to SharePoint Online and Microsoft Graph API.
#>

# ----------------------------------------------
# Set Variables
# ----------------------------------------------
$tenantname = "m365x61250205"                                   # This is your tenant name
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51"                # This is your Tenant ID
$searchRegion = "NAM"                                          # Region for Microsoft Graph search

# ----------------------------------------------
# Initialize Parameters - Do not change
# ----------------------------------------------
$sites = @()
$output = @()
$inputfile = $null
$outputfile = $null
$log = $null
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# ----------------------------------------------
# Input / Output and Log Files
# ----------------------------------------------
#$inputfile = 'C:\temp\sitelist.csv' #comment this line to run against all SPO Sites, otherwise use an input file.
$outputfile = "$env:TEMP\" + 'SPOSharingLinks' + $date + '_' + "incremental.csv"
$log = "$env:TEMP\" + 'SPOSharingLinks' + $date + '_' + "logfile.log"
# Initialize sharing links output file
$sharingLinksOutputFile = "$env:TEMP\" + 'SPO_SharingLinks_' + $date + '.csv'

# ----------------------------------------------
# Logging Function
# ----------------------------------------------
Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText
    )
    if ($LogName -ne $null) {
        # log the date and time in the text file along with the data passed
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : $LogEntryText" | Out-File -FilePath $LogName -append;
    }
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
            if ($_.Exception.Response -ne $null) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                
                if ($statusCode -eq 429 -or $statusCode -eq 503) {
                    $isThrottling = $true
                    
                    # Try to get the Retry-After header
                    $retryAfterHeader = $_.Exception.Response.Headers["Retry-After"]
                    
                    if ($retryAfterHeader) {
                        $waitTime = [int]$retryAfterHeader
                        Write-LogEntry -LogName $Log -LogEntryText "Throttling detected for $Operation. Retry-After header: $waitTime seconds."
                    }
                    else {
                        # Use exponential backoff if no Retry-After header
                        $waitTime = [Math]::Pow(2, $retryCount) * 10
                        Write-LogEntry -LogName $Log -LogEntryText "Throttling detected for $Operation. No Retry-After header. Using backoff: $waitTime seconds."
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
                    Write-LogEntry -LogName $Log -LogEntryText "PnP throttling detected for $Operation. Waiting for $waitTime seconds."
                }
                else {
                    # Use exponential backoff
                    $waitTime = [Math]::Pow(2, $retryCount) * 10
                    Write-LogEntry -LogName $Log -LogEntryText "PnP throttling detected for $Operation. Using backoff: $waitTime seconds."
                }
            }
            
            if ($isThrottling) {
                $retryCount++
                
                if ($retryCount -le $MaxRetries) {
                    Write-Host "  Throttling detected for $Operation. Retrying in $waitTime seconds... (Attempt $retryCount of $MaxRetries)" -ForegroundColor Yellow
                    Write-LogEntry -LogName $Log -LogEntryText "Waiting $waitTime seconds before retry #$retryCount for $Operation."
                    Start-Sleep -Seconds $waitTime
                    continue
                }
            }
            
            # If we reach here, it's either not throttling or we've exceeded retries
            Write-Host "Error in $Operation (Retry #$retryCount): $errorMessage" -ForegroundColor Red
            Write-LogEntry -LogName $Log -LogEntryText "Error in $Operation (Retry #$retryCount): $errorMessage"
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
    Write-LogEntry -LogName $Log -LogEntryText "Successfully connected to SharePoint Admin Center: $adminUrl"
}
catch {
    Write-Host "Error connecting to SharePoint Admin Center ($adminUrl): $_" -ForegroundColor Red
    Write-LogEntry -LogName $Log -LogEntryText "Error connecting to SharePoint Admin Center ($adminUrl): $_"
    exit
}

# ----------------------------------------------
# Get Site List
# ----------------------------------------------
if ($inputfile -and (Test-Path -Path $inputfile)) {
    try {
        $sites = Import-csv -path $inputfile -Header 'URL'
        Write-LogEntry -LogName $Log -LogEntryText "Using sites from input file: $inputfile"
        Write-Host "Reading sites from input file: $inputfile" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Error reading input file '$inputfile': $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error reading input file '$inputfile': $_"
        exit
    }
}
else {
    Write-Host "Getting site list from tenant (this might take a while)..." -ForegroundColor Yellow
    Write-LogEntry -LogName $Log -LogEntryText "Getting sites using Get-PnPTenantSite (no input file specified or found)"
    try {
        # Ensure we are connected to Admin Center before this call
        Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop
        $sites = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPTenantSite
        } -Operation "Get-PnPTenantSite"
        
        Write-Host "Found $($sites.Count) sites." -ForegroundColor Green
        Write-LogEntry -LogName $Log -LogEntryText "Retrieved $($sites.Count) sites using Get-PnPTenantSite."
    }
    catch {
        Write-Host "Error getting site list from tenant: $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error getting site list from tenant: $_"
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
$sharingLinksHeaders = "Site URL,Site Owner,IB Mode,IB Segment,Site Template,Sharing Group Name,Sharing Link Members,File URL,File Owner,IsTeamsConnected,SharingCapability,Last Content Modified"
Set-Content -Path $sharingLinksOutputFile -Value $sharingLinksHeaders
Write-LogEntry -LogName $Log -LogEntryText "Initialized sharing links output file: $sharingLinksOutputFile"

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
Function Process-SiteSharingLinks {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [object] $SiteData
    )
    
    # Check if this site has sharing links groups
    $sharingLinkGroups = $SiteData."SP Groups On Site" | Where-Object { $_ -like "SharingLinks*" }
    
    if ($sharingLinkGroups.Count -gt 0) {
        Write-Host "  Processing $($sharingLinkGroups.Count) sharing link groups for site: $SiteUrl" -ForegroundColor Yellow
        Write-LogEntry -LogName $Log -LogEntryText "Processing $($sharingLinkGroups.Count) sharing link groups for site: $SiteUrl"
        
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
                }
                
                # Write directly to the CSV file
                $csvLine | Export-Csv -Path $sharingLinksOutputFile -Append -NoTypeInformation -Force
                Write-LogEntry -LogName $Log -LogEntryText "  Wrote sharing link data for group: $sharingGroup"
            }
        }
    }
}

# ----------------------------------------------
# Main Processing Loop
# ----------------------------------------------
$totalSites = $sites.Count
$processedCount = 0
$sitesWithSharingLinksCount = 0

foreach ($site in $sites) {
    $processedCount++
    $siteUrl = $site.Url
    
    Write-Host "Processing site $processedCount/$totalSites : $siteUrl" -ForegroundColor Cyan
    Write-LogEntry -LogName $Log -LogEntryText "Processing site $processedCount/$totalSites : $siteUrl"

    try {
        # Get Site Properties using the Admin connection context
        Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop # Ensure admin context
        $siteprops = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPTenantSite -Identity $siteUrl
        } -Operation "Get-PnPTenantSite for $siteUrl" | 
        Select-Object Url, Owner, InformationBarrierMode, InformationBarrierSegments, 
        Template, SharingCapability, IsTeamsConnected, LastContentModifiedDate

        if ($null -eq $siteprops) { 
            Write-LogEntry -LogName $Log -LogEntryText "Failed to retrieve properties for site $siteUrl. Skipping."
            continue 
        }

        # Initialize site data with basic properties
        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops

        # Connect to the specific site for group/user information
        try {
            $currentPnPConnection = Connect-PnPOnline -Url $siteUrl @connectionParams -ErrorAction Stop
            Write-LogEntry -LogName $Log -LogEntryText "Successfully connected to site: $siteUrl"
            
            # SharePoint Group Processing
            $spGroups = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPGroup
            } -Operation "Get-PnPGroup for $siteUrl"
            
            Write-LogEntry -LogName $Log -LogEntryText "Found $($spGroups.Count) SP Groups on $siteUrl"
            
            ForEach ($spGroup in $spGroups) {
                if (!$spGroup -or !$spGroup.Title) { continue }
                
                $spGroupName = $spGroup.Title
                
                # Check if this is a sharing links group and add to collection
                if ($spGroupName -like "SharingLinks*") {
                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -SPGroupName $spGroupName
                    
                    # Get SP Group Members for sharing links groups only
                    if ($spGroup.Id) { 
                        $spGroupMembers = Invoke-WithThrottleHandling -ScriptBlock {
                            Get-PnPGroupMember -Identity $spGroup.Id
                        } -Operation "Get-PnPGroupMember for group $($spGroup.Title)"
                        
                        foreach ($member in $spGroupMembers) {
                            if (!$member -or !$member.LoginName) { continue }
                            
                            try {
                                $pnpUser = Invoke-WithThrottleHandling -ScriptBlock {
                                    Get-PnPUser -Identity $member.LoginName -ErrorAction SilentlyContinue
                                } -Operation "Get-PnPUser for $($member.LoginName)"
                                
                                if ($pnpUser) {
                                    $spUserName = $pnpUser.Title
                                    $spUserEmail = $pnpUser.Email
                                    
                                    # Call Update-SiteCollectionData for the specific user/group combo
                                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops `
                                        -AssociatedSPGroup $spGroupName -SPUserName $spUserName `
                                        -SPUserTitle $pnpUser.Title -SPUserEmail $spUserEmail
                                }
                                else {
                                    # Fallback to member title
                                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops `
                                        -AssociatedSPGroup $spGroupName -SPUserName $member.Title `
                                        -SPUserTitle $member.Title
                                }
                            }
                            catch { 
                                Write-LogEntry -LogName $Log -LogEntryText "Error processing member '$($member.LoginName)': ${_}"
                            }
                        }
                    }
                    
                    # Extract document ID and other sharing info from sharing link group name
                    try {
                        # Extract document ID if present in the sharing group name
                        if ($spGroupName -match "SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.([^.]+)") {
                            $documentId = $matches[1]
                            $sharingType = $matches[2]  # e.g. "Flexible", "View", "Edit"
                            
                            Write-LogEntry -LogName $Log -LogEntryText "Found document ID in sharing link: $documentId, Sharing Type: $sharingType"
                            
                            # Try to find the document using Microsoft Graph Search API
                            try {
                                # Get an access token for Microsoft Graph
                                $graphToken = Invoke-WithThrottleHandling -ScriptBlock {
                                    Get-PnPAccessToken
                                } -Operation "Get-PnPAccessToken"
                                
                                if ($graphToken) {
                                    # Prepare headers with the access token
                                    $headers = @{
                                        "Authorization" = "Bearer $graphToken"
                                        "Content-Type"  = "application/json"
                                    }
                                    
                                    # Prepare the search query as a PowerShell hashtable (easier to read and modify)
                                    $searchQuery = @{
                                        requests = @(
                                            @{
                                                entityTypes               = @("driveItem")
                                                query                     = @{
                                                    queryString = "UniqueID:$documentId"
                                                }
                                                from                      = $start
                                                size                      = $size
                                                sharePointOneDriveOptions = @{
                                                    includeContent = "sharedContent,privateContent"
                                                }
                                                region                    = $searchRegion
                                            }
                                        )
                                    }
                                    
                                    # Convert the hashtable to JSON
                                    $searchBody = $searchQuery | ConvertTo-Json -Depth 5
                                    
                                    # Execute the search query with throttling handling
                                    Write-LogEntry -LogName $Log -LogEntryText "Executing Microsoft Graph search for document ID: $documentId"
                                    $searchUrl = "https://graph.microsoft.com/v1.0/search/query"
                                    
                                    $searchResults = Invoke-WithThrottleHandling -ScriptBlock {
                                        Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body $searchBody
                                    } -Operation "Microsoft Graph Search API call for document ID $documentId"
                                    
                                    # Process search results
                                    $documentUrl = "Unable to locate document across tenant"
                                    $documentOwner = "Unknown - Document ID: $documentId"
                                    
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
                                            
                                            Write-LogEntry -LogName $Log -LogEntryText "Located document via Graph search: $documentUrl"
                                        }
                                    }
                                    else {
                                        Write-LogEntry -LogName $Log -LogEntryText "No matching documents found in Graph search for ID: $documentId"
                                    }
                                }
                                else {
                                    Write-LogEntry -LogName $Log -LogEntryText "Unable to get Graph access token for document search."
                                }
                            }
                            catch {
                                Write-LogEntry -LogName $Log -LogEntryText "Error searching for document via Graph API: ${_}"
                            }
                            
                            # Store the sharing link information 
                            if (-not $siteCollectionData[$siteUrl].ContainsKey("DocumentDetails")) {
                                $siteCollectionData[$siteUrl]["DocumentDetails"] = @{}
                            }
                            
                            $siteCollectionData[$siteUrl]["DocumentDetails"][$spGroupName] = @{
                                "DocumentId"    = $documentId
                                "SharingType"   = $sharingType
                                "DocumentUrl"   = $documentUrl
                                "DocumentOwner" = $documentOwner
                                "SharedOn"      = $siteUrl
                            }
                            
                            Write-LogEntry -LogName $Log -LogEntryText "Stored sharing information for document ID $documentId"
                        }
                    }
                    catch {
                        Write-LogEntry -LogName $Log -LogEntryText "Error extracting document ID from group name $($spGroupName) : ${_}"
                    }
                }
            }

            # Process and write sharing links data for this site immediately if any found
            if ($siteCollectionData[$siteUrl]["Has Sharing Links"]) {
                Process-SiteSharingLinks -SiteUrl $siteUrl -SiteData $siteCollectionData[$siteUrl]
                $sitesWithSharingLinksCount++
            }
        }
        catch { 
            Write-LogEntry -LogName $Log -LogEntryText "Could not connect to site $siteUrl : ${_}" 
        }
    }
    catch {
        Write-LogEntry -LogName $Log -LogEntryText "Error processing site $siteUrl : ${_}"
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
    Write-LogEntry -LogName $Log -LogEntryText "Total sites with sharing links: $sitesWithSharingLinksCount"
}
else {
    Write-Host "No site collections with sharing links found." -ForegroundColor Yellow
    Write-LogEntry -LogName $Log -LogEntryText "No site collections with sharing links found."
}

# ----------------------------------------------
# Disconnect and finish
# ----------------------------------------------
Disconnect-PnPOnline
Write-LogEntry -LogName $Log -LogEntryText "Script finished."
Write-Host "Script finished. Log file located at: $log" -ForegroundColor Green
