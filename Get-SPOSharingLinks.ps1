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
    File Name      : Get-SPOSharingLinks.ps1
    Author         : Mike Lee
    Date Created   : 5/5/25
    Prerequisite   : PnP PowerShell module, App registration in Entra ID with certificate

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
$inputfile = 'C:\temp\sitelist.csv'
$outputfile = "$env:TEMP\" + 'SPOSharingLinks' + $date + '_' + "output.csv"
$log = "$env:TEMP\" + 'SPOSharingLinks' + $date + '_' + "logfile.log"

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
        $sites = Get-PnPTenantSite # Excludes OneDrive by default
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
# Main Processing Loop
# ----------------------------------------------
$totalSites = $sites.Count
$processedCount = 0

foreach ($site in $sites) {
    $processedCount++
    $siteUrl = $site.Url
    
    Write-Host "Processing site $processedCount/$totalSites : $siteUrl" -ForegroundColor Cyan
    Write-LogEntry -LogName $Log -LogEntryText "Processing site $processedCount/$totalSites : $siteUrl"

    try {
        # Get Site Properties using the Admin connection context
        Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop # Ensure admin context
        $siteprops = Get-PnPTenantSite -Identity $siteUrl | 
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
            $spGroups = Get-PnPGroup
            Write-LogEntry -LogName $Log -LogEntryText "Found $($spGroups.Count) SP Groups on $siteUrl"
            
            ForEach ($spGroup in $spGroups) {
                if (!$spGroup -or !$spGroup.Title) { continue }
                
                $spGroupName = $spGroup.Title
                
                # Check if this is a sharing links group and add to collection
                if ($spGroupName -like "SharingLinks*") {
                    Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteprops -SPGroupName $spGroupName
                    
                    # Get SP Group Members for sharing links groups only
                    if ($spGroup.Id) { 
                        $spGroupMembers = Get-PnPGroupMember -Identity $spGroup.Id
                        
                        foreach ($member in $spGroupMembers) {
                            if (!$member -or !$member.LoginName) { continue }
                            
                            try {
                                $pnpUser = Get-PnPUser -Identity $member.LoginName -ErrorAction SilentlyContinue
                                
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
                                $graphToken = Get-PnPGraphAccessToken
                                
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
                                    
                                    # Execute the search query
                                    Write-LogEntry -LogName $Log -LogEntryText "Executing Microsoft Graph search for document ID: $documentId"
                                    $searchUrl = "https://graph.microsoft.com/v1.0/search/query"
                                    $searchResults = Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body $searchBody
                                    
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
$finalOutput = [System.Collections.Generic.List[PSObject]]::new()

foreach ($siteUrl in $siteCollectionData.Keys) {
    $siteData = $siteCollectionData[$siteUrl]
    
    # Create export item with only needed properties
    $exportItem = [PSCustomObject]@{
        URL                     = $siteData.URL
        Owner                   = $siteData.Owner
        "IB Mode"               = $siteData."IB Mode"
        "IB Segment"            = $siteData."IB Segment"
        Template                = $siteData.Template
        SharingCapability       = $siteData.SharingCapability
        IsTeamsConnected        = $siteData.IsTeamsConnected
        LastContentModifiedDate = $siteData.LastContentModifiedDate
        "Has Sharing Links"     = if ($siteData."Has Sharing Links") { "True" } else { "False" }
        "SP Groups On Site"     = ($siteData."SP Groups On Site" -join ';')
    }
    $finalOutput.Add($exportItem)
}

# ----------------------------------------------
# Output sharing links related data
# ----------------------------------------------
if ($finalOutput.Count -gt 0) {
    Write-Host "Processing sharing links data for output..." -ForegroundColor Green
    try {
        $sharingLinksOutput = [System.Collections.Generic.List[PSObject]]::new()
        
        # Filter to only sites with sharing links
        $sitesWithSharingLinks = $finalOutput | Where-Object { $_."Has Sharing Links" -eq "True" }
        
        if ($sitesWithSharingLinks.Count -gt 0) {
            Write-Host "Found $($sitesWithSharingLinks.Count) site collections with sharing links" -ForegroundColor Green
            
            foreach ($site in $sitesWithSharingLinks) {
                $siteUrl = $site.URL
                $siteData = $siteCollectionData[$siteUrl]
                
                # Get all sharing links groups from this site
                $sharingLinkGroups = $siteData."SP Groups On Site" | Where-Object { $_ -like "SharingLinks*" }
                
                foreach ($sharingGroup in $sharingLinkGroups) {
                    # Get users in this sharing links group
                    $groupMembers = $siteData."SP Users" | Where-Object { $_.AssociatedSPGroup -eq $sharingGroup }
                    
                    if ($groupMembers.Count -gt 0) {
                        # Format members as "Name <Email>"
                        $membersFormatted = ($groupMembers | ForEach-Object {
                                $emailStr = $_.Email | Out-String -NoNewline
                                "$($_.Name) <$emailStr>"
                            }) -join ';'
                        
                        # Get document details if available
                        $documentUrl = "Not found"
                        $documentOwner = "Not found"
                        if ($siteData.ContainsKey("DocumentDetails") -and $siteData["DocumentDetails"].ContainsKey($sharingGroup)) {
                            $documentUrl = $siteData["DocumentDetails"][$sharingGroup]["DocumentUrl"]
                            $documentOwner = $siteData["DocumentDetails"][$sharingGroup]["DocumentOwner"]
                        }
                        
                        # Create output item for this sharing link group
                        $sharingLinkItem = [PSCustomObject]@{
                            "Site URL"              = $site.URL
                            "Site Owner"            = $site.Owner
                            "IB Mode"               = $site."IB Mode"
                            "IB Segment"            = $site."IB Segment"
                            "Site Template"         = $site.Template
                            "Sharing Group Name"    = $sharingGroup
                            "Sharing Link Members"  = $membersFormatted
                            "File URL"              = $documentUrl
                            "File Owner"            = $documentOwner
                            "IsTeamsConnected"      = $site.IsTeamsConnected
                            "SharingCapability"     = $site.SharingCapability
                            "Last Content Modified" = $site.LastContentModifiedDate
                        }
                        
                        $sharingLinksOutput.Add($sharingLinkItem)
                    }
                }
            }
            
            if ($sharingLinksOutput.Count -gt 0) {
                $sharingLinksOutputFile = "$env:TEMP\" + 'SPO_SharingLinks_' + $date + '.csv'
                $sharingLinksOutput | Export-Csv -Path $sharingLinksOutputFile -NoTypeInformation -Encoding UTF8
                Write-Host "Sharing links data successfully written to: $sharingLinksOutputFile" -ForegroundColor Green
                Write-LogEntry -LogName $Log -LogEntryText "Sharing links data successfully written to: $sharingLinksOutputFile"
            }
            else {
                Write-Host "No detailed sharing links information found to export." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "No site collections with sharing links found." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error processing sharing links output: $($_)" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error processing sharing links output: $($_)"
    }
}

# ----------------------------------------------
# Disconnect and finish
# ----------------------------------------------
Disconnect-PnPOnline
Write-LogEntry -LogName $Log -LogEntryText "Script finished."
Write-Host "Script finished. Log file located at: $log" -ForegroundColor Green
