#Requires -Module PnP.PowerShell
<#
.SYNOPSIS
    Identifies Flexible sharing links containing external One-Time Passcode (OTP) users to assess the impact of
    MC1243549 - Retirement of SharePoint OTP and transition to Microsoft Entra B2B guest accounts.

.DESCRIPTION
    This script scans SharePoint Online sites to identify all Flexible sharing links that contain external users,
    specifically targeting One-Time Passcode (OTP) users as part of the MC1243549 retirement impact assessment.

    For each Flexible sharing link with external users, the script attempts to confirm whether those users are
    OTP users by looking them up in the site's User Information List. OTP users are identified by the
    "urn:spo:guest#" pattern in their login name, distinguishing them from proper Entra B2B guest accounts.

    Organization sharing links are excluded from this report as they are not in scope for OTP retirement.
    This script runs in Detection-only mode and makes NO modifications to sharing links or permissions.
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
    The region for Microsoft Graph search operations (e.g., "NAM", "EUR", "APC", "GBR", "CAN", etc.).
    Leave empty (default) to auto-detect the correct region for your tenant.

.PARAMETER Mode
    This script runs in Detection-only mode. No sharing links or permissions are modified.
    The Mode parameter is not used and can be omitted.

.PARAMETER debugLogging
    When set to $true, the script logs detailed DEBUG operations for troubleshooting.
    When set to $false, only INFO and ERROR operations are logged.

.PARAMETER inputfile
    Optional. Path to a CSV file containing a list of SharePoint site URLs (one URL per line, or with a "URL" header).
    If not specified, the script will process all sites in the tenant.

.PARAMETER GetOneDriveInfo
    When set to $true, the script scans OneDrive for Business (personal) sites ONLY.
    When set to $false (default), OneDrive sites are skipped and only SharePoint sites are scanned.

.OUTPUTS
    - CSV file containing Flexible sharing links that contain external OTP users (MC1243549 impact scope only)
    - Log file with operation details and errors
    
    Only links with confirmed external OTP users are written to the CSV.
    Flexible links with no external OTP users are excluded — they are not impacted by OTP retirement.
    
    CSV Output Columns:
    - Site URL: SharePoint site containing the sharing link
    - Site Owner: Owner of the SharePoint site
    - IB Mode: Information Barrier mode setting
    - IB Segment: Information Barrier segments
    - Site Template: SharePoint site template
    - Sharing Group Name: Name of the SharePoint sharing group
    - Sharing Link Members: Users who have access through the sharing link
    - File URL: Direct URL to the shared file or list item
    - File Owner: Owner/creator of the shared file
    - Filename: Name of the shared file or list item
    - SharingType: Type of sharing (Flexible only - Organization links are excluded from this report)
    - Sharing Link URL: Direct URL of the sharing link
    - Link Expiration Date: When the sharing link expires
    - IsTeamsConnected: Whether the site is connected to Microsoft Teams
    - SharingCapability: Site-level sharing capability setting
    - Last Content Modified: Last modification date of the site content
    - Search Status: Indicates if the document was found in search results
      * "Found" - Document located and indexed in search
      * "Found (REST Fallback)" - Document located via SharePoint REST API (not yet indexed in search)
      * "File Not Found" - Document not found via search, REST, or group description lookup;
        the sharing link likely points to a deleted or moved file (orphaned sharing group)
      * "Search Error" - An unexpected error occurred during all lookup attempts
      * "Not Searched" - Search was not attempted
    - Has External OTP Users: Whether the Flexible link contains external OTP users (True/False)
    - External OTP Users: Semicolon-separated list of external user emails confirmed as OTP users
      via the site's User Information List (identified by the "urn:spo:guest#" login pattern)
    - OTP Confirmed: Whether at least one external user was confirmed as an OTP user via the
      User Information List (True/False)

.NOTES
    Authors: Mike Lee
    Created: 3/24/2026
    Updated: 3/25/2026 - added multi-geo Graph Search region detection and handling
    Updated: 3/26/2026 - added support for scanning OneDrive Sites (optional parameter)
    Purpose: MC1243549 - Retirement of SharePoint One-Time Passcode (SPO OTP) and transition
             to Microsoft Entra B2B guest accounts. Run this script to identify Flexible sharing
             links that expose OTP users, so admins can assess the retirement impact.

    - Requires PnP.PowerShell 2.x or above
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
    # Scan all sites in the tenant and report on Flexible sharing links with OTP users
    .\SPO-OTPHelper.ps1

.EXAMPLE
    # Scan a specific list of sites from a CSV file
    $inputfile = "C:\temp\sites.csv"
    .\SPO-OTPHelper.ps1
#>

# ----------------------------------------------
# Set Variables
# ----------------------------------------------
$tenantname = "m365cpi13246019"                                   # This is your tenant name
$appID = "abc64618-283f-47ba-a185-50d935d51d57"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"                # This is your Tenant ID
$searchRegion = ""                                              # Region for Microsoft Graph search (leave empty to auto-detect, or set explicitly: US/NAM/EUR/APC/GBR/CAN/IND/AUS/JPN/DEU/etc.)
$debugLogging = $false                                         # Set to $true for detailed DEBUG logging, $false for INFO and ERROR logging only
$GetOneDriveInfo = $false                                      # Set to $true to scan OneDrive sites ONLY; $false (default) scans SharePoint sites and skips OneDrive

# ----------------------------------------------
# Initialize Parameters - Do not change
# ----------------------------------------------
$sites = @()
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# ----------------------------------------------
# Input / Output and Log Files
# ----------------------------------------------
$inputfile = "" #If no input file specified, will process all sites in the tenant
$log = "$env:TEMP\" + 'SPOSharingLinks' + $date + '_' + "logfile.log"
# Initialize sharing links output file
$sharingLinksOutputFile = "$env:TEMP\" + 'SPO_OTP_Impact_' + $date + '.csv'

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
# PnP Version Detection and Graph Token Helper Function
# ----------------------------------------------
Function Get-PnPGraphTokenCompatible {
    <#
    .SYNOPSIS
    Gets a Graph access token using the appropriate command based on PnP PowerShell version.
    
    .DESCRIPTION
    Automatically detects PnP PowerShell version and uses:
    - Get-PnPAccessToken for PnP PowerShell 3.0+
    - Get-PnPGraphAccessToken for PnP PowerShell 2.x and earlier
    #>
    
    try {
        # Get the PnP PowerShell module version
        $pnpModule = Get-Module -Name "PnP.PowerShell" -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        
        if (-not $pnpModule) {
            throw "PnP.PowerShell module not found"
        }
        
        $majorVersion = $pnpModule.Version.Major
        Write-DebugLog -LogName $Log -LogEntryText "Detected PnP.PowerShell version: $($pnpModule.Version) (Major: $majorVersion)"
        
        if ($majorVersion -ge 3) {
            # PnP PowerShell 3.0+ uses Get-PnPAccessToken
            Write-DebugLog -LogName $Log -LogEntryText "Using Get-PnPAccessToken for PnP PowerShell 3.0+"
            return Get-PnPAccessToken
        }
        else {
            # PnP PowerShell 2.x and earlier uses Get-PnPGraphAccessToken
            Write-DebugLog -LogName $Log -LogEntryText "Using Get-PnPGraphAccessToken for PnP PowerShell 2.x"
            return Get-PnPGraphAccessToken
        }
    }
    catch {
        # Fallback: try the newer command first, then the older one
        Write-DebugLog -LogName $Log -LogEntryText "Version detection failed, trying fallback approach: $_"
        
        try {
            Write-DebugLog -LogName $Log -LogEntryText "Fallback: Attempting Get-PnPAccessToken (PnP 3.0+)"
            return Get-PnPAccessToken
        }
        catch {
            Write-DebugLog -LogName $Log -LogEntryText "Fallback: Attempting Get-PnPGraphAccessToken (PnP 2.x)"
            return Get-PnPGraphAccessToken
        }
    }
}

# ----------------------------------------------
# Script Configuration - OTP Retirement Impact Assessment (MC1243549)
# This script runs in Detection-only mode; no sharing links or permissions are modified.
# ----------------------------------------------

$scriptMode = "DETECTION"

Write-Host "Script is running in DETECTION mode (MC1243549 - OTP Retirement Impact Assessment)" -ForegroundColor Cyan
Write-InfoLog -LogName $Log -LogEntryText "MC1243549 OTP Retirement Impact Assessment - Running in DETECTION mode. Scanning Flexible sharing links for external OTP users. No modifications will be made."


# ----------------------------------------------
# Connection Parameters
# ----------------------------------------------
Add-Type -AssemblyName System.Web   # Required for [System.Web.HttpUtility]::UrlDecode used in filename extraction
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
# Graph Search Region Detection
# ----------------------------------------------
Function Test-GraphSearchRegion {
    <#
    .SYNOPSIS
    Tests whether a given region code is valid for this tenant's Graph Search API.
    Returns $true if the region works (HTTP 200), $false if it fails (HTTP 400 = wrong region).
    #>
    param(
        [string] $Region,
        [hashtable] $Headers
    )

    $testQuery = @{
        requests = @(
            @{
                entityTypes               = @("driveItem")
                query                     = @{ queryString = "test" }
                from                      = 0
                size                      = 1
                sharePointOneDriveOptions = @{ includeContent = "sharedContent" }
                region                    = $Region
            }
        )
    }

    try {
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/search/query" `
            -Headers $Headers -Method Post `
            -Body ($testQuery | ConvertTo-Json -Depth 5) `
            -ErrorAction Stop | Out-Null
        Write-DebugLog -LogName $Log -LogEntryText "Region probe succeeded: $Region"
        return $true
    }
    catch {
        Write-DebugLog -LogName $Log -LogEntryText "Region probe failed for '$Region': $_"
        return $false
    }
}

Function Get-GraphSearchRegion {
    <#
    .SYNOPSIS
    Auto-detects the correct Microsoft Graph Search region for this tenant by probing the
    search API with candidate regions until one succeeds (HTTP 200 vs 400 Bad Request).
    Regions are tried in order of Microsoft 365 global customer volume.
    #>

    Write-Host "  Auto-detecting Microsoft Graph Search region..." -ForegroundColor Cyan
    Write-InfoLog -LogName $Log -LogEntryText "Auto-detecting Microsoft Graph Search region (\$searchRegion is empty)"

    try {
        $graphToken = Get-PnPGraphTokenCompatible
        if (-not $graphToken) {
            Write-ErrorLog -LogName $Log -LogEntryText "Unable to obtain Graph token for region auto-detection, defaulting to NAM"
            return "NAM"
        }

        $headers = @{
            "Authorization" = "Bearer $graphToken"
            "Content-Type"  = "application/json"
        }

        # Ordered by Microsoft 365 global market share / likelihood
        # "US" is an alternate code for NAM used in some non-multi-geo tenants
        $regionsToProbe = @(
            "NAM", "US", "EUR", "APC", "GBR", "CAN",
            "IND", "AUS", "JPN", "DEU", "ZAF",
            "ARE", "CHE", "NOR", "KOR", "SWE",
            "TWN", "FRA", "ITA", "MEX", "LAM",
            "NZL", "SGP", "BRA", "MYS", "QAT", "POL"
        )

        foreach ($region in $regionsToProbe) {
            if (Test-GraphSearchRegion -Region $region -Headers $headers) {
                Write-Host "  Graph Search region auto-detected: $region" -ForegroundColor Green
                Write-InfoLog -LogName $Log -LogEntryText "Graph Search region auto-detected: $region"
                return $region
            }
        }
    }
    catch {
        Write-ErrorLog -LogName $Log -LogEntryText "Unexpected error during Graph Search region auto-detection: $_"
    }

    Write-Host "  Could not auto-detect Graph Search region, defaulting to NAM" -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "Graph Search region auto-detection failed, defaulting to NAM"
    return "NAM"
}

# ----------------------------------------------
# Multi-Geo: Resolve the correct Graph Search region for a specific site's geo location.
# Get-PnPTenantSite returns a GeoLocation property (e.g. "NAM", "EUR", "APC") per site.
# In multi-geo tenants, sites in satellite geos must use the matching region code.
# Results are cached to avoid re-probing the same geo multiple times.
# For non-multi-geo tenants, GeoLocation is empty and the global $searchRegion is used.
# ----------------------------------------------
Function Get-SiteSearchRegion {
    param(
        [string] $GeoLocation
    )

    # Non-multi-geo: no geo location set — use the already-resolved global region
    if ([string]::IsNullOrWhiteSpace($GeoLocation)) {
        return $searchRegion
    }

    $geoKey = $GeoLocation.ToUpper()

    # Cache hit — return immediately without re-probing
    if ($geoRegionCache.ContainsKey($geoKey)) {
        Write-DebugLog -LogName $Log -LogEntryText "Geo region cache hit for '$geoKey': $($geoRegionCache[$geoKey])"
        return $geoRegionCache[$geoKey]
    }

    Write-Host "  Probing Graph Search region for geo location: $geoKey" -ForegroundColor Cyan
    Write-InfoLog -LogName $Log -LogEntryText "Multi-geo: probing Graph Search region for geo '$geoKey'"

    try {
        $graphToken = Get-PnPGraphTokenCompatible
        if (-not $graphToken) {
            Write-ErrorLog -LogName $Log -LogEntryText "Unable to obtain Graph token for geo region probe '$geoKey', using default: $searchRegion"
            $geoRegionCache[$geoKey] = $searchRegion
            return $searchRegion
        }

        $headers = @{
            "Authorization" = "Bearer $graphToken"
            "Content-Type"  = "application/json"
        }

        # Try the GeoLocation code directly — SPO geo codes match Graph Search region codes
        if (Test-GraphSearchRegion -Region $geoKey -Headers $headers) {
            Write-Host "  Graph Search region confirmed for geo '$geoKey': $geoKey" -ForegroundColor Green
            Write-InfoLog -LogName $Log -LogEntryText "Multi-geo: Graph Search region confirmed for geo '$geoKey': $geoKey"
            $geoRegionCache[$geoKey] = $geoKey
            return $geoKey
        }

        # "US" is sometimes used instead of "NAM" for the North America geo
        if ($geoKey -eq "NAM" -and (Test-GraphSearchRegion -Region "US" -Headers $headers)) {
            Write-Host "  Graph Search region confirmed for geo 'NAM': US (alternate code)" -ForegroundColor Green
            Write-InfoLog -LogName $Log -LogEntryText "Multi-geo: Graph Search region confirmed for geo 'NAM' using alternate code 'US'"
            $geoRegionCache[$geoKey] = "US"
            return "US"
        }
    }
    catch {
        Write-ErrorLog -LogName $Log -LogEntryText "Error probing Graph Search region for geo '$geoKey': $_"
    }

    # Fallback: use the globally resolved region
    Write-Host "  Could not confirm Graph Search region for geo '$geoKey', using default: $searchRegion" -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "Multi-geo: could not confirm region for geo '$geoKey', falling back to: $searchRegion"
    $geoRegionCache[$geoKey] = $searchRegion
    return $searchRegion
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
# Resolve Graph Search Region
# ----------------------------------------------
if ([string]::IsNullOrWhiteSpace($searchRegion)) {
    $searchRegion = Get-GraphSearchRegion
}
else {
    Write-Host "  Using configured Graph Search region: $searchRegion" -ForegroundColor Cyan
    Write-InfoLog -LogName $Log -LogEntryText "Using configured Graph Search region: $searchRegion"
}

# ----------------------------------------------
# Get Site List
# ----------------------------------------------
if ($inputfile -and (Test-Path -Path $inputfile)) {
    Write-Host "Processing input file: $inputfile" -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "Processing input file: $inputfile"
    
    try {
        $firstLine = Get-Content -Path $inputfile -TotalCount 1 -ErrorAction Stop
        if ($firstLine -match '^URL$|^url$|^Url$') {
            # File already has a URL header row — import without adding a duplicate header
            $sites = Import-Csv -Path $inputfile
        }
        else {
            # No header row — treat first column as URL
            $sites = Import-Csv -Path $inputfile -Header 'URL'
        }
        Write-Host "Found $($sites.Count) sites in input file." -ForegroundColor Green
        Write-InfoLog -LogName $Log -LogEntryText "Loaded $($sites.Count) sites from input file: $inputfile"
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
        if ($GetOneDriveInfo) {
            Write-Host "  Mode: OneDrive sites ONLY ($GetOneDriveInfo = $true)" -ForegroundColor Cyan
            Write-InfoLog -LogName $Log -LogEntryText "GetOneDriveInfo=$true : retrieving OneDrive personal sites only"
            $sites = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPTenantSite -IncludeOneDriveSites:$true | Where-Object {
                    $_.Url -like "*-my.sharepoint.com/personal/*" -and
                    $_.Status -eq "Active" -and
                    $_.ArchiveStatus -eq "NotArchived" -and
                    $_.SharingCapability -ne "Disabled" -and
                    -not [string]::IsNullOrEmpty($_.Url)
                }
            } -Operation "Get-PnPTenantSite (OneDrive sites only)"
        }
        else {
            Write-Host "  Mode: SharePoint sites only (OneDrive excluded)" -ForegroundColor Cyan
            Write-InfoLog -LogName $Log -LogEntryText "GetOneDriveInfo=$false : retrieving SharePoint sites, skipping OneDrive"
            $sites = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPTenantSite -IncludeOneDriveSites:$false | Where-Object {
                    $_.Template -notmatch "SRCHCEN|MYSITE|APPCATALOG|PWS|POINTPUBLISHINGTOPIC|SPSMSITEHOST|EHS|REVIEWCTR|TENANTADMIN" -and
                    $_.Status -eq "Active" -and
                    $_.ArchiveStatus -eq "NotArchived" -and
                    $_.SharingCapability -ne "Disabled" -and
                    -not [string]::IsNullOrEmpty($_.Url)
                }
            } -Operation "Get-PnPTenantSite with optimized filtering"
        }
        
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

# Cache for geo location -> validated Graph Search region (avoids re-probing the same geo)
$geoRegionCache = @{}

# ----------------------------------------------
# Initialize the sharing links output file with headers
# ----------------------------------------------
$sharingLinksHeaders = "Site URL,Site Owner,IB Mode,IB Segment,Site Template,Sharing Group Name,Sharing Link Members,File URL,File Owner,Filename,SharingType,Sharing Link URL,Link Expiration Date,IsTeamsConnected,SharingCapability,Last Content Modified,Search Status,Has External OTP Users,External OTP Users,OTP Confirmed"
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
            "GeoLocation"             = $SiteProperties.GeoLocation
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
        
        # Debug: Log when we store sharing group members
        if ($AssociatedSPGroup -like "SharingLinks*") {
            Write-DebugLog -LogName $Log -LogEntryText "STORED user in site data: Group='$AssociatedSPGroup', Name='$SPUserName', Title='$SPUserTitle', Email='$SPUserEmail'"
        }
    }
    else {
        # Debug: Log when we skip storing a user
        if ($AssociatedSPGroup -like "SharingLinks*") {
            Write-DebugLog -LogName $Log -LogEntryText "SKIPPED storing user: SPUserName empty: $([string]::IsNullOrWhiteSpace($SPUserName)), AssociatedSPGroup empty: $([string]::IsNullOrWhiteSpace($AssociatedSPGroup)), Values: Name='$SPUserName', Group='$AssociatedSPGroup'"
        }
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
            # Get document details first — $searchStatus is needed by $membersFormatted below
            $documentUrl = "Not found"
            $documentOwner = "Not found"
            $documentItemType = "Not found"
            $sharingLinkUrl = "Not found"
            $linkExpirationDate = "Not found"
            $searchStatus = "Not Searched"
            if ($SiteData.ContainsKey("DocumentDetails") -and $SiteData["DocumentDetails"].ContainsKey($sharingGroup)) {
                $documentUrl = $SiteData["DocumentDetails"][$sharingGroup]["DocumentUrl"]
                $documentOwner = $SiteData["DocumentDetails"][$sharingGroup]["DocumentOwner"]
                $documentItemType = $SiteData["DocumentDetails"][$sharingGroup]["DocumentItemType"]
                $sharingLinkUrl = $SiteData["DocumentDetails"][$sharingGroup]["SharingLinkUrl"]
                $linkExpirationDate = $SiteData["DocumentDetails"][$sharingGroup]["ExpirationDate"]
                $searchStatus = $SiteData["DocumentDetails"][$sharingGroup]["SearchStatus"]
                Write-DebugLog -LogName $Log -LogEntryText "Retrieved document details for $sharingGroup - URL: $documentUrl, Owner: $documentOwner, Type: $documentItemType, LinkURL: $sharingLinkUrl, Expiration: $linkExpirationDate, SearchStatus: $searchStatus"
            }
            else {
                Write-DebugLog -LogName $Log -LogEntryText "No document details found for sharing group: $sharingGroup. DocumentDetails exists: $($SiteData.ContainsKey('DocumentDetails')), Group key exists: $(if ($SiteData.ContainsKey('DocumentDetails')) { $SiteData['DocumentDetails'].ContainsKey($sharingGroup) } else { 'N/A' })"
            }

            # Get users in this sharing links group
            $groupMembers = $SiteData."SP Users" | Where-Object { $_.AssociatedSPGroup -eq $sharingGroup }
            
            # Debug: Log what members we found for this sharing group
            Write-DebugLog -LogName $Log -LogEntryText "Processing sharing group '$sharingGroup' - found $($groupMembers.Count) members in site data"
            
            # Debug: Also show ALL users for this site to verify data storage
            $allSiteUsers = $SiteData."SP Users"
            $allSharingUsers = $allSiteUsers | Where-Object { $_.AssociatedSPGroup -like "SharingLinks*" }
            Write-DebugLog -LogName $Log -LogEntryText "Site has $($allSiteUsers.Count) total users, $($allSharingUsers.Count) in sharing groups"
            
            if ($groupMembers.Count -gt 0) {
                foreach ($member in $groupMembers) {
                    Write-DebugLog -LogName $Log -LogEntryText "  Found member: Name='$($member.Name)', Title='$($member.Title)', Email='$($member.Email)'"
                }
            }
            else {
                # Debug: If no members found for this specific group, check if there are any users with similar group names
                $similarGroups = $allSiteUsers | Where-Object { $_.AssociatedSPGroup -like "*$($sharingGroup.Split('.')[1])*" }
                Write-DebugLog -LogName $Log -LogEntryText "  No members found for exact group name '$sharingGroup'. Found $($similarGroups.Count) users with similar group patterns."
                foreach ($similarUser in $similarGroups) {
                    Write-DebugLog -LogName $Log -LogEntryText "    Similar: Group='$($similarUser.AssociatedSPGroup)', Name='$($similarUser.Name)'"
                }
            }
            
            # Format members as "Name <Email>" - handle empty groups based on search status
            $membersFormatted = if ($groupMembers.Count -gt 0) {
                ($groupMembers | ForEach-Object {
                    $emailStr = if ($_.Email) { $_.Email | Out-String -NoNewline } else { "" }
                    "$($_.Name) <$emailStr>"
                }) -join ';'
            }
            else {
                # Check if the file was not locatable to provide better context for empty member lists
                if ($searchStatus -eq "File Not Found") {
                    "File Not Found"
                }
                elseif ($searchStatus -eq "Search Error") {
                    "Search Error"
                }
                else {
                    "No members"
                }
            }
            
            # Extract filename from the document URL
            $filename = "Not found"
            if ($documentUrl -ne "Not found" -and $documentUrl -ne "File Not Found" -and $documentUrl -ne "Search Error" -and -not [string]::IsNullOrWhiteSpace($documentUrl)) {
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
            elseif ($documentUrl -eq "File Not Found") {
                $filename = "File Not Found"
            }
            elseif ($documentUrl -eq "Search Error") {
                $filename = "Search Error"
            }
            
            # Determine sharing type - this script focuses on Flexible links for OTP retirement assessment
            $sharingType = "Unknown"
            if ($sharingGroup -like "*Flexible*") {
                $sharingType = "Flexible"
            }
            elseif ($sharingGroup -like "*Organization*") {
                # Organization links are out of scope for OTP retirement assessment (MC1243549)
                Write-DebugLog -LogName $Log -LogEntryText "  Skipping Organization link (out of scope for OTP assessment): $sharingGroup"
                continue
            }

            # Retrieve OTP detection results for this Flexible link group
            $hasExternalOTPUsers = "False"
            $externalOTPUsersList = ""
            $otpConfirmed = "False"

            if ($SiteData.ContainsKey("OTP Detection") -and $SiteData["OTP Detection"].ContainsKey($sharingGroup)) {
                $otpData = $SiteData["OTP Detection"][$sharingGroup]

                if ($otpData.HasExternalOTPUsers) {
                    $hasExternalOTPUsers = "True"
                    $otpUserStrings = [System.Collections.Generic.List[string]]::new()
                    $anyConfirmed = $false

                    foreach ($otpUser in $otpData.ExternalOTPUsers) {
                        $userLabel = if (-not [string]::IsNullOrWhiteSpace($otpUser.Email)) { $otpUser.Email } else { $otpUser.LoginName }
                        if ($otpUser.ConfirmedInUIL) {
                            $anyConfirmed = $true
                        }
                        $otpUserStrings.Add($userLabel)
                    }

                    $externalOTPUsersList = $otpUserStrings -join "; "
                    $otpConfirmed = if ($anyConfirmed) { "True" } else { "False" }
                }
            }

            # Skip links with no external OTP users — they are not impacted by MC1243549 OTP retirement
            if ($hasExternalOTPUsers -ne "True") {
                Write-DebugLog -LogName $Log -LogEntryText "  Skipping group (no OTP users): $sharingGroup"
                continue
            }

            # Create CSV line
            $csvLine = [PSCustomObject]@{
                "Site URL"               = $SiteData.URL
                "Site Owner"             = $SiteData.Owner
                "IB Mode"                = $SiteData."IB Mode"
                "IB Segment"             = $SiteData."IB Segment"
                "Site Template"          = $SiteData.Template
                "Sharing Group Name"     = $sharingGroup
                "Sharing Link Members"   = $membersFormatted
                "File URL"               = $documentUrl
                "File Owner"             = $documentOwner
                "Filename"               = $filename
                "SharingType"            = $sharingType
                "Sharing Link URL"       = $sharingLinkUrl
                "Link Expiration Date"   = $linkExpirationDate
                "IsTeamsConnected"       = $SiteData.IsTeamsConnected
                "SharingCapability"      = $SiteData.SharingCapability
                "Last Content Modified"  = $SiteData.LastContentModifiedDate
                "Search Status"          = $searchStatus
                "Has External OTP Users" = $hasExternalOTPUsers
                "External OTP Users"     = $externalOTPUsersList
                "OTP Confirmed"          = $otpConfirmed
            }

            # Write directly to the CSV file
            $csvLine | Export-Csv -Path $sharingLinksOutputFile -Append -NoTypeInformation -Force
            Write-DebugLog -LogName $Log -LogEntryText "  Wrote sharing link data for group: $sharingGroup"
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
            Get-PnPGraphTokenCompatible
        } -Operation "Get Graph access token (version-compatible) for $LogContext"

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

        # Try driveItem search first (document library files)
        $driveItemSearchQuery = @{
            requests = @(
                @{
                    entityTypes               = @("driveItem")
                    query                     = @{ queryString = "`"$DocumentId`"" }
                    from                      = 0
                    size                      = 25
                    sharePointOneDriveOptions = @{ includeContent = "sharedContent,privateContent" }
                    region                    = $SearchRegion
                }
            )
        }

        $searchResults = Invoke-WithThrottleHandling -ScriptBlock {
            Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body ($driveItemSearchQuery | ConvertTo-Json -Depth 5)
        } -Operation "$LogContext - Graph Search driveItem for $DocumentId"

        if ($searchResults.value -and
            $searchResults.value[0].hitsContainers -and
            $searchResults.value[0].hitsContainers[0].hits -and
            $searchResults.value[0].hitsContainers[0].hits.Count -gt 0) {

            $resource = $searchResults.value[0].hitsContainers[0].hits[0].resource
            if ($resource) {
                $result.Found = $true
                $result.ItemType = "driveItem"
                if ($resource.webUrl) { $result.DocumentUrl = $resource.webUrl }
                if ($resource.createdBy.user.displayName) {
                    $result.DocumentOwner = if ($resource.createdBy.user.email) {
                        "$($resource.createdBy.user.displayName) <$($resource.createdBy.user.email)>"
                    }
                    else { $resource.createdBy.user.displayName }
                }
                Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Located via driveItem: $($result.DocumentUrl)"
                $itemFound = $true
            }
        }

        # Fallback to listItem search
        if (-not $itemFound) {
            $listItemSearchQuery = @{
                requests = @(
                    @{
                        entityTypes               = @("listItem")
                        query                     = @{ queryString = "`"$DocumentId`"" }
                        from                      = 0
                        size                      = 25
                        sharePointOneDriveOptions = @{ includeContent = "sharedContent,privateContent" }
                        region                    = $SearchRegion
                    }
                )
            }

            $listItemResults = Invoke-WithThrottleHandling -ScriptBlock {
                Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body ($listItemSearchQuery | ConvertTo-Json -Depth 5)
            } -Operation "$LogContext - Graph Search listItem for $DocumentId"

            if ($listItemResults.value -and
                $listItemResults.value[0].hitsContainers -and
                $listItemResults.value[0].hitsContainers[0].hits -and
                $listItemResults.value[0].hitsContainers[0].hits.Count -gt 0) {

                $resource = $listItemResults.value[0].hitsContainers[0].hits[0].resource
                if ($resource) {
                    $result.Found = $true
                    $result.ItemType = "listItem"
                    if ($resource.webUrl) { $result.DocumentUrl = $resource.webUrl }
                    if ($resource.createdBy.user.displayName) {
                        $result.DocumentOwner = if ($resource.createdBy.user.email) {
                            "$($resource.createdBy.user.displayName) <$($resource.createdBy.user.email)>"
                        }
                        else { $resource.createdBy.user.displayName }
                    }
                    Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Located via listItem: $($result.DocumentUrl)"
                }
            }
            else {
                Write-DebugLog -LogName $Log -LogEntryText "$LogContext - No results found for ID: $DocumentId"
            }
        }
    }
    catch {
        Write-ErrorLog -LogName $Log -LogEntryText "$LogContext - Error searching for document via Graph API: $_"
    }

    return $result
}

# ----------------------------------------------
# Function to retrieve document properties directly via SharePoint REST API
# Used as a fallback when Graph search has not yet indexed the document.
# Only the document UniqueId (GUID from the sharing group name) is required.
# Tries GetFileById first (document library files), then GetListItemByUniqueId (list items/pages).
# ----------------------------------------------
Function Get-DocumentPropertiesByUniqueId {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [Parameter(Mandatory = $true)]
        [string] $DocumentId,
        [Parameter(Mandatory = $false)]
        [string] $GroupDescription = "",
        [Parameter(Mandatory = $false)]
        [string] $LogContext = "Document lookup"
    )

    $result = @{
        Found         = $false
        DocumentUrl   = ""
        DocumentOwner = ""
        ItemType      = ""
    }

    try {
        Write-DebugLog -LogName $Log -LogEntryText "$LogContext - REST fallback for document ID: $DocumentId on $SiteUrl"

        # Extract tenant root URL (e.g. https://contoso.sharepoint.com) for full URL construction
        $uri = [System.Uri]$SiteUrl
        $tenantRoot = "$($uri.Scheme)://$($uri.Host)"

        # --- Attempt 1: GetFileById (document library files) ---
        # Note: Called directly (not via Invoke-WithThrottleHandling) because a 404/FileNotFound
        # response is expected and normal when the item is a list item rather than a file.
        try {
            $getFileUrl = "/_api/web/GetFileById('$DocumentId')?`$select=ServerRelativeUrl,Author/Title,Author/Email&`$expand=Author"
            $fileResponse = Invoke-PnPSPRestMethod -Method Get -Url $getFileUrl -ErrorAction Stop

            if ($fileResponse -and -not [string]::IsNullOrWhiteSpace($fileResponse.ServerRelativeUrl)) {
                $result.Found = $true
                $result.ItemType = "driveItem"
                $result.DocumentUrl = $tenantRoot + $fileResponse.ServerRelativeUrl

                if ($fileResponse.Author) {
                    $authorEmail = $fileResponse.Author.Email
                    $result.DocumentOwner = if (-not [string]::IsNullOrWhiteSpace($authorEmail)) {
                        "$($fileResponse.Author.Title) <$authorEmail>"
                    }
                    else { $fileResponse.Author.Title }
                }
                Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Found via GetFileById REST: $($result.DocumentUrl)"
                return $result
            }
        }
        catch {
            Write-DebugLog -LogName $Log -LogEntryText "$LogContext - GetFileById failed (may be a list item): $_"
        }

        # --- Attempt 2: SharePoint REST search API (site-scoped) ---
        # Graph search can be blocked by tenant search-restriction settings (per the admin banner
        # "Your organization's admin has restricted Search from accessing certain SharePoint sites").
        # The SharePoint REST search endpoint runs under the PnP certificate connection and is not
        # subject to the same Graph-level restrictions, so it can reach items that Graph missed.
        # This also handles list items that GetListItemByUniqueId doesn't support.
        #
        # NOTE on response shape: Invoke-PnPSPRestMethod returns OData verbose JSON, where every
        # array is wrapped in { "results": [...] }.  We must unwrap via .results before iterating.
        try {
            # Use a plain quoted keyword search for the GUID — the UniqueID managed property is
            # reliable for document library files but may not be indexed for generic list items.
            # A phrase search for the GUID string reliably matches whichever field carries it.
            $spSearchUrl = "/_api/search/query?querytext='%22$DocumentId%22'&SelectProperties='Title,Path,DefaultEncodingURL,AuthorOWSUSER,FileExtension,UniqueID'&RowLimit=5&TrimDuplicates=false"
            $spSearchResponse = Invoke-PnPSPRestMethod -Method Get -Url $spSearchUrl -ErrorAction Stop

            # Unwrap the OData verbose { results: [] } wrapper that PnP returns for all arrays
            $searchRows = $null
            $primaryResult = $spSearchResponse.PrimaryQueryResult
            if ($primaryResult) {
                $relevantResults = $primaryResult.RelevantResults
                if ($relevantResults) {
                    $table = $relevantResults.Table
                    if ($table) {
                        $rawRows = $table.Rows
                        # Handle both direct array and OData-verbose { results: [] } wrapper
                        $searchRows = if ($rawRows.results) { $rawRows.results } else { $rawRows }
                    }
                }
            }

            if ($searchRows -and $searchRows.Count -gt 0) {
                # Unwrap Cells the same way
                $rawCells = $searchRows[0].Cells
                $cells = if ($rawCells.results) { $rawCells.results } else { $rawCells }

                $itemPath = ($cells | Where-Object { $_.Key -eq "Path" } | Select-Object -First 1).Value
                $itemDefaultEncodingUrl = ($cells | Where-Object { $_.Key -eq "DefaultEncodingURL" } | Select-Object -First 1).Value
                $authorField = ($cells | Where-Object { $_.Key -eq "AuthorOWSUSER" } | Select-Object -First 1).Value

                Write-DebugLog -LogName $Log -LogEntryText "$LogContext - SP REST search returned $($searchRows.Count) row(s). Path='$itemPath', EncodingURL='$itemDefaultEncodingUrl'"

                # Prefer DefaultEncodingURL (direct item URL) over Path when available
                $resolvedUrl = if (-not [string]::IsNullOrWhiteSpace($itemDefaultEncodingUrl)) {
                    $itemDefaultEncodingUrl
                }
                elseif (-not [string]::IsNullOrWhiteSpace($itemPath)) {
                    $itemPath
                }
                else { "" }

                if (-not [string]::IsNullOrWhiteSpace($resolvedUrl)) {
                    $result.Found = $true
                    $result.ItemType = "listItem"
                    $result.DocumentUrl = $resolvedUrl

                    # AuthorOWSUSER format: "id | Display Name | email" — extract display name and email
                    if (-not [string]::IsNullOrWhiteSpace($authorField)) {
                        $authorParts = $authorField -split '\s*\|\s*'
                        $authorName = if ($authorParts.Count -ge 2) { $authorParts[1].Trim() } else { $authorField }
                        $authorEmail = if ($authorParts.Count -ge 3) { $authorParts[2].Trim() } else { "" }
                        $result.DocumentOwner = if (-not [string]::IsNullOrWhiteSpace($authorEmail)) {
                            "$authorName <$authorEmail>"
                        }
                        else { $authorName }
                    }
                    Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Found via SP REST search: $resolvedUrl"
                    return $result
                }
            }
            else {
                Write-DebugLog -LogName $Log -LogEntryText "$LogContext - SP REST search returned no rows for ID: $DocumentId"
            }
        }
        catch {
            Write-DebugLog -LogName $Log -LogEntryText "$LogContext - SP REST search failed: $_"
        }

        # --- Attempt 3: GetListItemByUniqueId (list items / wiki pages) ---
        # Note: Called directly (not via Invoke-WithThrottleHandling) because a ResourceNotFoundException
        # is expected and normal when the item genuinely does not exist on this site.
        try {
            $getListItemUrl = "/_api/web/GetListItemByUniqueId('$DocumentId')?`$select=EncodedAbsUrl,FileRef,Author/Title,Author/EMail&`$expand=Author"
            $listItemResponse = Invoke-PnPSPRestMethod -Method Get -Url $getListItemUrl -ErrorAction Stop

            if ($listItemResponse) {
                $itemUrl = if (-not [string]::IsNullOrWhiteSpace($listItemResponse.EncodedAbsUrl)) {
                    $listItemResponse.EncodedAbsUrl
                }
                elseif (-not [string]::IsNullOrWhiteSpace($listItemResponse.FileRef)) {
                    $tenantRoot + $listItemResponse.FileRef
                }
                else { "" }

                if (-not [string]::IsNullOrWhiteSpace($itemUrl)) {
                    $result.Found = $true
                    $result.ItemType = "listItem"
                    $result.DocumentUrl = $itemUrl

                    if ($listItemResponse.Author) {
                        $authorEmail = $listItemResponse.Author.EMail
                        $result.DocumentOwner = if (-not [string]::IsNullOrWhiteSpace($authorEmail)) {
                            "$($listItemResponse.Author.Title) <$authorEmail>"
                        }
                        else { $listItemResponse.Author.Title }
                    }
                    Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Found via GetListItemByUniqueId REST: $($result.DocumentUrl)"
                    return $result
                }
            }
        }
        catch {
            Write-DebugLog -LogName $Log -LogEntryText "$LogContext - GetListItemByUniqueId failed: $_"
        }

        Write-DebugLog -LogName $Log -LogEntryText "$LogContext - UniqueId-based lookups exhausted; trying group description path for ID: $DocumentId"

        # --- Attempt 4: Use the SharingLinks group Description as a server-relative path ---
        # SharePoint sets the Description of SharingLinks.* groups to the server-relative path of
        # the shared item. This works even when the search index is stale, as long as the file exists.
        if (-not [string]::IsNullOrWhiteSpace($GroupDescription)) {
            $descPath = $GroupDescription.Trim()
            Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Trying group description as path: '$descPath'"

            # Normalise: strip any absolute prefix so we always have a server-relative path
            if ($descPath -match "^https?://[^/]+(/.+)$") { $descPath = $matches[1] }

            if ($descPath -match "^/") {
                try {
                    # GetFileByServerRelativePath handles spaces and special characters correctly
                    $encodedPath = $descPath.Replace("'", "''")  # escape single quotes for OData
                    $byPathUrl = "/_api/web/GetFileByServerRelativePath(decodedurl='$encodedPath')?`$select=ServerRelativeUrl,Author/Title,Author/Email&`$expand=Author"
                    $pathResponse = Invoke-PnPSPRestMethod -Method Get -Url $byPathUrl -ErrorAction Stop

                    if ($pathResponse -and -not [string]::IsNullOrWhiteSpace($pathResponse.ServerRelativeUrl)) {
                        $result.Found = $true
                        $result.ItemType = "driveItem"
                        $result.DocumentUrl = $tenantRoot + $pathResponse.ServerRelativeUrl

                        if ($pathResponse.Author) {
                            $authorEmail = $pathResponse.Author.Email
                            $result.DocumentOwner = if (-not [string]::IsNullOrWhiteSpace($authorEmail)) {
                                "$($pathResponse.Author.Title) <$authorEmail>"
                            }
                            else { $pathResponse.Author.Title }
                        }
                        Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Found via group Description path: $($result.DocumentUrl)"
                        return $result
                    }
                }
                catch {
                    Write-DebugLog -LogName $Log -LogEntryText "$LogContext - Group description path lookup failed: $_"
                }
            }
        }

        Write-DebugLog -LogName $Log -LogEntryText "$LogContext - All fallbacks exhausted; document not found for ID: $DocumentId"
    }
    catch {
        Write-ErrorLog -LogName $Log -LogEntryText "$LogContext - Unexpected error in REST fallback for document ID $DocumentId : $_"
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
        $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if (-not $currentConnection -or $currentConnection.Url -ne $SiteUrl) {
            Connect-PnPOnline -Url $SiteUrl @connectionParams -ErrorAction Stop
        }

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

            if ($groupName -match "SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.") {
                $documentId = $matches[1]
                Write-DebugLog -LogName $Log -LogEntryText "Processing sharing group: $groupName with document ID: $documentId"

                if ($siteCollectionData[$SiteUrl].ContainsKey("DocumentDetails") -and
                    $siteCollectionData[$SiteUrl]["DocumentDetails"].ContainsKey($groupName) -and
                    -not [string]::IsNullOrWhiteSpace($siteCollectionData[$SiteUrl]["DocumentDetails"][$groupName]["DocumentUrl"])) {

                    $docUrl = $siteCollectionData[$SiteUrl]["DocumentDetails"][$groupName]["DocumentUrl"]

                    try {
                        $sharingLinks = $null
                        $sharingLinkUrl = "Not found"
                        $expirationDate = "No expiration"

                        $sharingLinks = Invoke-WithThrottleHandling -ScriptBlock {
                            Get-PnPFileSharingLink -Identity $documentId -ErrorAction SilentlyContinue
                        } -Operation "Get sharing links for document ID: $documentId"

                        if ($sharingLinks -and $sharingLinks.Count -gt 0) {
                            Write-DebugLog -LogName $Log -LogEntryText "Found $($sharingLinks.Count) sharing links for document"

                            $matchingLink = $sharingLinks | Where-Object { $_.Id -and $groupName -like "*$($_.Id)*" } | Select-Object -First 1

                            if ($matchingLink) {
                                $sharingLinkUrl = if ($matchingLink.link -and $matchingLink.link.WebUrl) {
                                    $matchingLink.link.WebUrl
                                }
                                else { "Not found" }

                                # Populate members from GrantedToIdentitiesV2 / GrantedToV2
                                $grantedIdentities = if ($matchingLink.GrantedToIdentitiesV2 -and $matchingLink.GrantedToIdentitiesV2.Count -gt 0) {
                                    $matchingLink.GrantedToIdentitiesV2 | ForEach-Object { $_.User }
                                }
                                elseif ($matchingLink.GrantedToV2 -and $matchingLink.GrantedToV2.Count -gt 0) {
                                    $matchingLink.GrantedToV2 | ForEach-Object { $_.User }
                                }
                                else { @() }

                                foreach ($u in $grantedIdentities) {
                                    if (-not $u) { continue }
                                    $memberEmail = if ($u.Email) { $u.Email }       else { "" }
                                    $memberDisplayName = if ($u.DisplayName) { $u.DisplayName } else { $memberEmail }
                                    $memberLoginName = if ($u.Id) { $u.Id }          else { $memberEmail }

                                    $existingMember = $siteCollectionData[$SiteUrl]["SP Users"] | Where-Object {
                                        $_.AssociatedSPGroup -eq $groupName -and
                                        ($_.Name -eq $memberLoginName -or $_.Email -eq $memberEmail)
                                    }
                                    if (-not $existingMember) {
                                        $siteCollectionData[$SiteUrl]["SP Users"].Add([PSCustomObject]@{
                                                AssociatedSPGroup = $groupName
                                                Name              = $memberLoginName
                                                Title             = $memberDisplayName
                                                Email             = $memberEmail
                                            })
                                    }
                                }

                                # Get expiration date
                                $rawExp = if ($matchingLink.link -and $matchingLink.link.ExpirationDateTime) {
                                    $matchingLink.link.ExpirationDateTime
                                }
                                elseif ($matchingLink.ExpirationDateTime) {
                                    $matchingLink.ExpirationDateTime
                                }
                                else { $null }

                                if ($rawExp) {
                                    try { $expirationDate = ([DateTime]::Parse($rawExp)).ToString("yyyy-MM-dd HH:mm:ss") }
                                    catch { $expirationDate = $rawExp }
                                }

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

# ----------------------------------------------
# Function to detect external OTP users in Flexible sharing links and confirm via User Information List
# MC1243549 - Retirement of SharePoint One-Time Passcode (SPO OTP)
# OTP users are identified by "urn:spo:guest#" in their SharePoint login name.
# ----------------------------------------------
Function Find-FlexibleLinkOTPUsers {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl
    )

    Write-Host "  Scanning for external OTP users in Flexible sharing links: $SiteUrl" -ForegroundColor Cyan
    Write-InfoLog -LogName $Log -LogEntryText "MC1243549: Scanning for external OTP users in Flexible sharing links on site: $SiteUrl"

    try {
        Connect-PnPOnline -Url $SiteUrl @connectionParams -ErrorAction Stop

        $flexibleGroups = $siteCollectionData[$SiteUrl]["SP Groups On Site"] | Where-Object { $_ -like "SharingLinks*Flexible*" }

        if ($flexibleGroups.Count -eq 0) {
            Write-DebugLog -LogName $Log -LogEntryText "No Flexible sharing groups found for site: $SiteUrl"
            return
        }

        if (-not $siteCollectionData[$SiteUrl].ContainsKey("OTP Detection")) {
            $siteCollectionData[$SiteUrl]["OTP Detection"] = @{}
        }

        # Always query the User Information List (UIL) first.
        # We cannot rely solely on login-name pattern matching for the pre-scan because
        # OTP users added via Get-SharingLinkUrls (GrantedToIdentitiesV2) are stored with
        # their Entra object ID (a GUID) as the login name — not the urn:spo:guest# form.
        # Querying UIL first lets us catch those users via email matching in the per-member loop.
        Write-InfoLog -LogName $Log -LogEntryText "Querying User Information List for OTP user confirmation on site: $SiteUrl"

        $uiListOTPUsers = @{}
        try {
            $siteUsersList = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPUser -ErrorAction SilentlyContinue
            } -Operation "Get site users (User Information List) for OTP confirmation on $SiteUrl"

            if ($siteUsersList) {
                foreach ($siteUser in $siteUsersList) {
                    $ulLoginName = $siteUser.LoginName
                    if ($ulLoginName -match "urn:spo:guest#|urn%3aspo%3aguest#") {
                        $email = $siteUser.Email
                        if ([string]::IsNullOrWhiteSpace($email)) {
                            if ($ulLoginName -match "urn:spo:guest#(.+)$") {
                                $email = $matches[1].Trim()
                            }
                            elseif ($ulLoginName -match "urn%3aspo%3aguest#(.+)$") {
                                $email = [System.Uri]::UnescapeDataString($matches[1].Trim())
                            }
                        }
                        if (-not [string]::IsNullOrWhiteSpace($email)) {
                            $uiListOTPUsers[$email.ToLower()] = @{
                                LoginName   = $ulLoginName
                                Email       = $email
                                DisplayName = $siteUser.Title
                            }
                            Write-DebugLog -LogName $Log -LogEntryText "UIL OTP user found: LoginName='$ulLoginName', Email='$email'"
                        }
                    }
                }
                Write-InfoLog -LogName $Log -LogEntryText "User Information List contains $($uiListOTPUsers.Count) OTP user(s) for site: $SiteUrl"
            }
        }
        catch {
            Write-ErrorLog -LogName $Log -LogEntryText "Error querying User Information List for site $SiteUrl : $_"
        }

        # Early-return optimisation: skip per-group work if the UIL has no OTP users
        # AND no member has the urn:spo:guest# pattern in their login name.
        if ($uiListOTPUsers.Count -eq 0) {
            $hasAnyOTPLoginPattern = $false
            foreach ($groupName in $flexibleGroups) {
                $members = $siteCollectionData[$SiteUrl]["SP Users"] | Where-Object { $_.AssociatedSPGroup -eq $groupName }
                foreach ($member in $members) {
                    if ([System.Uri]::UnescapeDataString($member.Name) -match "urn:spo:guest#") {
                        $hasAnyOTPLoginPattern = $true
                        break
                    }
                }
                if ($hasAnyOTPLoginPattern) { break }
            }

            if (-not $hasAnyOTPLoginPattern) {
                Write-DebugLog -LogName $Log -LogEntryText "No OTP users in UIL and no urn:spo:guest# login patterns found for site: $SiteUrl"
                foreach ($groupName in $flexibleGroups) {
                    $siteCollectionData[$SiteUrl]["OTP Detection"][$groupName] = @{
                        HasExternalOTPUsers = $false
                        ExternalOTPUsers    = @()
                    }
                }
                return
            }
        }
        else {
            Write-Host "    UIL contains $($uiListOTPUsers.Count) OTP user(s) - checking Flexible sharing group membership" -ForegroundColor Yellow
        }

        foreach ($groupName in $flexibleGroups) {
            $members = $siteCollectionData[$SiteUrl]["SP Users"] | Where-Object { $_.AssociatedSPGroup -eq $groupName }
            $externalOTPUsersInGroup = [System.Collections.Generic.List[PSObject]]::new()
            $hasExternalOTPUsers = $false

            foreach ($member in $members) {
                $loginName = $member.Name
                $email = $member.Email
                # URL-decode the login name to handle all encoding variants before matching.
                $decodedLoginName = [System.Uri]::UnescapeDataString($loginName)
                $isOTPPattern = $decodedLoginName -match "urn:spo:guest#"

                if ($isOTPPattern -and [string]::IsNullOrWhiteSpace($email)) {
                    # $decodedLoginName is already URL-decoded; extract email after the # token
                    if ($decodedLoginName -match "urn:spo:guest#(.+)$") {
                        $email = $matches[1].Trim()
                    }
                }

                # Only flag users whose login name contains the OTP guest token
                # OR whose email is confirmed as an OTP user in the UIL.
                # The second check catches OTP users stored with a GUID login name
                # (added via Get-SharingLinkUrls from GrantedToIdentitiesV2, where
                # Graph returns the Entra object ID rather than the SPO login name).
                $isConfirmedByUIL = (
                    -not [string]::IsNullOrWhiteSpace($email) -and
                    $uiListOTPUsers.ContainsKey($email.ToLower())
                )

                if ($isOTPPattern -or $isConfirmedByUIL) {
                    $hasExternalOTPUsers = $true
                    # $isConfirmedByUIL already computed above; use it directly
                    $confirmedInUIL = $isConfirmedByUIL
                    $uilDetails = ""

                    if ($isConfirmedByUIL) {
                        $uilDetails = $uiListOTPUsers[$email.ToLower()].LoginName
                        Write-DebugLog -LogName $Log -LogEntryText "OTP confirmed via UIL for '$groupName': Email='$email', UIL LoginName='$uilDetails'"
                    }
                    else {
                        Write-DebugLog -LogName $Log -LogEntryText "OTP login pattern found but not confirmed in UIL for '$groupName': Email='$email', LoginName='$loginName'"
                    }

                    $externalOTPUsersInGroup.Add([PSCustomObject]@{
                            LoginName      = $loginName
                            Email          = $email
                            DisplayName    = $member.Title
                            IsOTPPattern   = $isOTPPattern
                            ConfirmedInUIL = $confirmedInUIL
                            UILDetails     = $uilDetails
                        })
                }
            }

            $siteCollectionData[$SiteUrl]["OTP Detection"][$groupName] = @{
                HasExternalOTPUsers = $hasExternalOTPUsers
                ExternalOTPUsers    = $externalOTPUsersInGroup
            }

            if ($hasExternalOTPUsers) {
                Write-InfoLog -LogName $Log -LogEntryText "Flexible group '$groupName' contains $($externalOTPUsersInGroup.Count) external OTP user(s)"
            }
        }
    }
    catch {
        Write-Host "  Error detecting OTP users for site $SiteUrl : $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error detecting OTP users for site $SiteUrl : $_"
    }
}

# ----------------------------------------------
# Main Processing Loop
# ----------------------------------------------
$totalSites = $sites.Count
$processedCount = 0
$sitesWithSharingLinksCount = 0
$sitesWithOTPUsersCount = 0

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "SCRIPT MODE: DETECTION (MC1243549 - OTP Retirement Impact Assessment)" -ForegroundColor Cyan
Write-Host "  - Scanning Flexible sharing links for external OTP users" -ForegroundColor Cyan
Write-Host "  - Confirming OTP users via each site's User Information List" -ForegroundColor Cyan
Write-Host "  - Organization sharing links are EXCLUDED from this report" -ForegroundColor Cyan
Write-Host "  - NO modifications will be made to permissions or sharing links" -ForegroundColor Cyan
Write-Host "  - Results will be saved to: $sharingLinksOutputFile" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""
Write-InfoLog -LogName $Log -LogEntryText "Starting to process $totalSites sites in $scriptMode mode"

foreach ($site in $sites) {
    $processedCount++
    $siteUrl = ""

    if ($site.URL) { $siteUrl = $site.URL }
    elseif ($site.Url) { $siteUrl = $site.Url }
    else { $siteUrl = $site.ToString() }

    if ([string]::IsNullOrWhiteSpace($siteUrl)) { continue }

    Write-Host "Processing site $processedCount of $totalSites : $siteUrl" -ForegroundColor Green
    Write-InfoLog -LogName $Log -LogEntryText "Processing site $processedCount of $totalSites : $siteUrl"

    try {
        try {
            Connect-PnPOnline -Url $siteUrl @connectionParams -ErrorAction Stop

            # Get site properties via Admin connection
            Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop
            $siteProperties = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPTenantSite -Identity $siteUrl
            } -Operation "Get site properties for $siteUrl"

            # Reconnect to the site for group processing
            Connect-PnPOnline -Url $siteUrl @connectionParams -ErrorAction Stop

            # Resolve the Graph Search region for this site's geo location (multi-geo aware).
            # For non-multi-geo tenants, GeoLocation is empty and Get-SiteSearchRegion returns $searchRegion.
            $siteSearchRegion = Get-SiteSearchRegion -GeoLocation $siteProperties.GeoLocation
            Write-DebugLog -LogName $Log -LogEntryText "Site '$siteUrl' geo='$($siteProperties.GeoLocation)' -> search region='$siteSearchRegion'"

            # Skip archived sites — safety-net for sites loaded from an input CSV file,
            # which cannot be pre-filtered. ArchiveStatus "NotArchived" is the only processable state.
            if ($siteProperties.ArchiveStatus -ne "NotArchived") {
                Write-Host "  Skipping archived site (ArchiveStatus=$($siteProperties.ArchiveStatus)): $siteUrl" -ForegroundColor DarkYellow
                Write-InfoLog -LogName $Log -LogEntryText "Skipping archived site (ArchiveStatus=$($siteProperties.ArchiveStatus)): $siteUrl"
                continue
            }

            Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteProperties

            $spGroups = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPGroup -Includes Description
            } -Operation "Get groups for site $siteUrl"

            foreach ($spGroup in $spGroups) {
                $spGroupName = $spGroup.Title

                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteProperties -SPGroupName $spGroupName

                $spUsers = Invoke-WithThrottleHandling -ScriptBlock {
                    $standardUsers = Get-PnPGroupMember -Identity $spGroup.Id -ErrorAction SilentlyContinue

                    if ($spGroupName -like "SharingLinks*") {
                        try {
                            $ctx = Get-PnPContext
                            $group = $ctx.Web.SiteGroups.GetById($spGroup.Id)
                            $users = $group.Users
                            $ctx.Load($users)
                            $ctx.ExecuteQuery()

                            $csomUsers = @()
                            foreach ($user in $users) {
                                $csomUsers += [PSCustomObject]@{
                                    Id            = $user.Id
                                    LoginName     = $user.LoginName
                                    Title         = $user.Title
                                    Email         = $user.Email
                                    PrincipalType = $user.PrincipalType
                                }
                            }
                            $allUsers = @($standardUsers) + @($csomUsers) | Group-Object LoginName | ForEach-Object { $_.Group[0] }
                            Write-DebugLog -LogName $Log -LogEntryText "Group '$spGroupName': Standard=$($standardUsers.Count), CSOM=$($csomUsers.Count), Combined=$($allUsers.Count)"
                            return $allUsers
                        }
                        catch {
                            Write-DebugLog -LogName $Log -LogEntryText "CSOM fallback failed for '$spGroupName': $_. Using standard results."
                            return $standardUsers
                        }
                    }
                    else {
                        return $standardUsers
                    }
                } -Operation "Get members for group $spGroupName"

                if ($spGroupName -like "SharingLinks*") {
                    Write-DebugLog -LogName $Log -LogEntryText "Sharing group '$spGroupName' has $($spUsers.Count) members"
                }

                foreach ($spUser in $spUsers) {
                    $hasValidLoginName = -not [string]::IsNullOrWhiteSpace($spUser.LoginName)
                    $hasValidId = $spUser.Id -ne $null -and $spUser.Id -gt 0

                    if ($spUser -and ($hasValidLoginName -or $hasValidId)) {
                        $userIdentifier = if (-not [string]::IsNullOrWhiteSpace($spUser.LoginName)) {
                            $spUser.LoginName
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace($spUser.Title)) {
                            $spUser.Title
                        }
                        else {
                            "User_$($spUser.Id)"
                        }
                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteProperties -AssociatedSPGroup $spGroupName -SPUserName $userIdentifier -SPUserTitle $spUser.Title -SPUserEmail $spUser.Email
                    }
                }

                # Extract document information from sharing groups
                if ($spGroupName -like "SharingLinks*") {
                    try {
                        if ($spGroupName -match "SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.") {
                            $documentId = $matches[1]
                            $sharingType = "Unknown"
                            $documentUrl = ""
                            $documentOwner = ""
                            $documentItemType = ""
                            $searchStatus = "Not Searched"

                            if ($spGroupName -like "*OrganizationView*") { $sharingType = "OrganizationView" }
                            elseif ($spGroupName -like "*OrganizationEdit*") { $sharingType = "OrganizationEdit" }
                            elseif ($spGroupName -like "*AnonymousAccess*") { $sharingType = "AnonymousAccess" }

                            try {
                                $graphToken = Invoke-WithThrottleHandling -ScriptBlock {
                                    Get-PnPGraphTokenCompatible
                                } -Operation "Get Graph access token for document search"

                                if ($graphToken) {
                                    $searchResult = Search-DocumentViaGraphAPI -DocumentId $documentId -SearchRegion $siteSearchRegion -LogContext "Main loop - document search"

                                    if ($searchResult.Found) {
                                        $searchStatus = "Found"
                                        $documentUrl = $searchResult.DocumentUrl
                                        $documentOwner = $searchResult.DocumentOwner
                                        $documentItemType = $searchResult.ItemType
                                    }
                                    else {
                                        # Document not in search index - try SharePoint REST API fallback using the UniqueId
                                        Write-DebugLog -LogName $Log -LogEntryText "Document not in search index - attempting REST fallback for ID: $documentId"
                                        $restResult = Get-DocumentPropertiesByUniqueId -SiteUrl $siteUrl -DocumentId $documentId -GroupDescription $spGroup.Description -LogContext "Main loop - REST fallback"
                                        if ($restResult.Found) {
                                            $searchStatus = "Found (REST Fallback)"
                                            $documentUrl = $restResult.DocumentUrl
                                            $documentOwner = $restResult.DocumentOwner
                                            $documentItemType = $restResult.ItemType
                                        }
                                        else {
                                            $searchStatus = "File Not Found"
                                            $documentUrl = "File Not Found"
                                            $documentOwner = "File Not Found"
                                            $documentItemType = "File Not Found"
                                        }
                                    }
                                }
                                else {
                                    Write-ErrorLog -LogName $Log -LogEntryText "Unable to get Graph access token for document search."
                                    # Graph token unavailable - REST fallback uses PnP certificate auth, not Graph token
                                    $restResult = Get-DocumentPropertiesByUniqueId -SiteUrl $siteUrl -DocumentId $documentId -GroupDescription $spGroup.Description -LogContext "Main loop - REST fallback (no Graph token)"
                                    if ($restResult.Found) {
                                        $searchStatus = "Found (REST Fallback)"
                                        $documentUrl = $restResult.DocumentUrl
                                        $documentOwner = $restResult.DocumentOwner
                                        $documentItemType = $restResult.ItemType
                                    }
                                    else {
                                        $searchStatus = "Search Error"
                                        $documentUrl = "Search Error"
                                        $documentOwner = "Search Error"
                                        $documentItemType = "Search Error"
                                    }
                                }
                            }
                            catch {
                                Write-ErrorLog -LogName $Log -LogEntryText "Error searching for document via Graph API: ${_}"
                                # Graph search threw an exception - try REST fallback
                                $restResult = Get-DocumentPropertiesByUniqueId -SiteUrl $siteUrl -DocumentId $documentId -GroupDescription $spGroup.Description -LogContext "Main loop - REST fallback after search error"
                                if ($restResult.Found) {
                                    $searchStatus = "Found (REST Fallback)"
                                    $documentUrl = $restResult.DocumentUrl
                                    $documentOwner = $restResult.DocumentOwner
                                    $documentItemType = $restResult.ItemType
                                }
                                else {
                                    $searchStatus = "Search Error"
                                    $documentUrl = "Search Error"
                                    $documentOwner = "Search Error"
                                    $documentItemType = "Search Error"
                                }
                            }

                            if (-not $siteCollectionData[$siteUrl].ContainsKey("DocumentDetails")) {
                                $siteCollectionData[$siteUrl]["DocumentDetails"] = @{}
                            }

                            $siteCollectionData[$siteUrl]["DocumentDetails"][$spGroupName] = @{
                                "DocumentId"       = $documentId
                                "SharingType"      = $sharingType
                                "DocumentUrl"      = $documentUrl
                                "DocumentOwner"    = $documentOwner
                                "DocumentItemType" = $documentItemType
                                "SearchStatus"     = $searchStatus
                                "SharedOn"         = $siteUrl
                                "SharingLinkUrl"   = ""
                                "ExpirationDate"   = ""
                            }
                        }
                    }
                    catch {
                        Write-ErrorLog -LogName $Log -LogEntryText "Error extracting document ID from group name $($spGroupName) : ${_}"
                    }
                }
            }

            # Process and write sharing links data for this site if any found
            if ($siteCollectionData[$siteUrl]["Has Sharing Links"]) {
                $sitesWithSharingLinksCount++

                # Collect sharing link URLs
                Get-SharingLinkUrls -SiteUrl $siteUrl

                # Detect external OTP users in Flexible sharing links (MC1243549 assessment)
                Find-FlexibleLinkOTPUsers -SiteUrl $siteUrl

                # Track sites with confirmed external OTP users for summary reporting
                if ($siteCollectionData[$siteUrl].ContainsKey("OTP Detection")) {
                    $siteHasOTP = $siteCollectionData[$siteUrl]["OTP Detection"].Values |
                    Where-Object { $_ -is [hashtable] -and $_.HasExternalOTPUsers -eq $true }
                    if ($siteHasOTP) { $sitesWithOTPUsersCount++ }
                }

                # Write sharing links data for this site (Flexible only)
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

# MC1243549 - OTP Retirement Impact Assessment Results
if ($sitesWithSharingLinksCount -gt 0) {
    Write-Host "Found $sitesWithSharingLinksCount site collection(s) with Flexible sharing links" -ForegroundColor Green
    Write-Host "Sites with external OTP users identified: $sitesWithOTPUsersCount" -ForegroundColor $(if ($sitesWithOTPUsersCount -gt 0) { "Yellow" } else { "Green" })
    Write-Host ""
    Write-Host "MC1243549 OTP Retirement Impact Assessment complete." -ForegroundColor Cyan
    Write-Host "Review the output CSV for Flexible links containing external OTP users:" -ForegroundColor Cyan
    Write-Host "  $sharingLinksOutputFile" -ForegroundColor Cyan
    Write-InfoLog -LogName $Log -LogEntryText "MC1243549 Assessment complete. Sites with Flexible sharing links: $sitesWithSharingLinksCount. Sites with external OTP users: $sitesWithOTPUsersCount. Output: $sharingLinksOutputFile"
}
else {
    Write-Host "No site collections with Flexible sharing links found." -ForegroundColor Green
    Write-InfoLog -LogName $Log -LogEntryText "MC1243549 Assessment: No site collections with Flexible sharing links found."
}

# ----------------------------------------------
# Disconnect and finish
# ----------------------------------------------
Disconnect-PnPOnline
Write-InfoLog -LogName $Log -LogEntryText "Script finished."
Write-Host "Script finished. Log file located at: $log" -ForegroundColor Green
