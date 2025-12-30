<#
.SYNOPSIS
    Identifies and processes SharePoint Online sharing links across the tenant.

.DESCRIPTION
    This script scans SharePoint Online sites to identify all sharing links, with a focus on Organization sharing links.
    It can optionally convert Organization sharing links to direct permissions and clean up corrupted Organization sharing groups.
    Flexible sharing links are detected and reported but never modified in remediation mode.
    The script supports scanning all sites in a tenant or a specific list of sites from a CSV file.

.PARAMETER tenantName
    The name of your Microsoft 365 tenant (without .onmicrosoft.com).

.PARAMETER appID
    The Entra (Azure AD) application ID used for authentication.

.PARAMETER thumbprint
    The certificate thumbprint for authentication.

.PARAMETER tenantId
    The tenant ID (GUID) for your Microsoft 365 tenant.

.PARAMETER inputFile
    Optional. Path to a CSV file containing either:
    1. A simple list of SharePoint site URLs (one URL per line or with "URL" header)
    2. The output CSV from a previous run of this script in report mode

    When using the script's own CSV output as input, the script will:
    - Only process sites that have Organization sharing links (identified by group names containing "Organization")
    - Automatically set Mode to "Remediation" for focused remediation
    - Skip other types of sharing links for focused remediation

    If not specified, the script will process all sites in the tenant.

.PARAMETER Mode
    Sets the script operation mode:
    - "Detection": Only inventories sharing links without making any modifications (report mode)
    - "Remediation": Converts Organization sharing links to direct permissions and removes Organization sharing groups (remediation mode). Flexible sharing links are not modified.
    Default: "Detection"

.PARAMETER ignoreFlexibleLinkGroups
    When set to $true, the script will ignore groups where the Group Type is 'Flexible'
    Flexible links are those which are direct sharing with another user.
    Reduces discovery time when sharing is widely used.

.PARAMETER removeAnyoneLinks
    When set to $true, and Mode = 'Remediation', Anonymous/Anyone links will be removed.
    No permissions for Anyone links are retained.

.PARAMETER cleanupCorruptedSharingGroups
    When set to $true, the script attempts to clean up empty or corrupted Organization sharing groups.
    When set to $false, no cleanup of sharing groups is performed.
    Note: Flexible sharing groups are never affected by cleanup operations.
    Note: This is automatically set to $true when Mode is set to "Remediation".

.PARAMETER logFilePath
    Output log file path. Optional
    Specify either the full path to a log file. If omitted then a log file will be created called 'SPO_SharingLinks_yyyyMMdd_HHmmss.log'

.PARAMETER outputFilePath
    Output results CSV path. Optional
    Specify either the full path to a csv file. If omitted then a csv file will be created called 'SPO_SharingLinks_yyyyMMdd_HHmmss.csv'

.PARAMETER debugLogging
    When set to $true, the script logs detailed DEBUG operations for troubleshooting.
    When set to $false, only INFO and ERROR operations are logged.

.OUTPUTS
    - CSV file containing detailed information about sharing links found, including search status for each document
    - Log file with operation details and errors

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
    - SharingType: Type of sharing (Organization, Flexible, etc.)
    - Sharing Link URL: Direct URL of the sharing link
    - Link Expiration Date: When the sharing link expires
    - IsTeamsConnected: Whether the site is connected to Microsoft Teams
    - SharingCapability: Site-level sharing capability setting
    - Last Content Modified: Last modification date of the site content
    - Search Status: Indicates if the document was found in search results
      * "Found" - Document located and indexed in search
      * "Not Found in Search" - Document exists but not indexed/searchable
      * "Search Error" - Error occurred during search operation
      * "Not Searched" - Search was not attempted
    - Link Removed: Whether the sharing link was removed (in remediation mode)

.NOTES
    Author         : Mike Lee
    Date Created   : 8/28/2025
    Update History :
        8/28/2025  - Initial script creation (Mike Lee)
        8/29/2025  - Updates (Mike Lee)
        12/29/2025 - Get-PnpGraphTokenCompatible caches discovery to reduce processing time (Craig Tolley)
        12/29/2025 - Use ArrayLists for performance improvements (Craig Tolley)
        12/29/2025 - Add in DocumentLastModified property to output (Craig Tolley)
        12/29/2025 - Minimise reconnects to improve performance (Craig Tolley)
        12/29/2025 - Update identification of Anonymous link types (Craig Tolley)
        12/29/2025 - Add option to ignore Flexible groups to reduce discovery time (Craig Tolley)
        12/29/2025 - Support removal of Anyone links (Craig Tolley)
        12/29/2025 - Dynamically look up site region to support multi-geo tenants (Craig Tolley)
        12/29/2025 - Move to parameters and update help (Craig Tolley)

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
    # Process all sites, only inventory sharing links without modifications (report mode)
    . .\Get-and-Remove-SPOSharingLinks.ps1 -tenantName m365cpi13246019 -appId abc64618-283f-47ba-a185-50d935d51d57 -thumbprint B696FDCFE1453F3FBC6031F54DE988DA0ED905A9 -tenantId 9cfc42cb-51da-4055-87e9-b20a170b6ba3

.EXAMPLE
    # Two-step process: Report then Remediate
    # Step 1: Run in report mode to generate CSV output
    . .\Get-and-Remove-SPOSharingLinks.ps1 -tenantName m365cpi13246019 -appId abc64618-283f-47ba-a185-50d935d51d57 -thumbprint B696FDCFE1453F3FBC6031F54DE988DA0ED905A9 -tenantId 9cfc42cb-51da-4055-87e9-b20a170b6ba3

    # Step 2: Use the generated CSV to remediate only Organization links
    $inputFile = ".\SPO_SharingLinks_2025-07-01_14-30-15.csv"
    # Note: Mode will be automatically set to "Remediation" when using script's CSV output
    . .\Get-and-Remove-SPOSharingLinks.ps1 -inputFile $inputFile -tenantName m365cpi13246019 -appId abc64618-283f-47ba-a185-50d935d51d57 -thumbprint B696FDCFE1453F3FBC6031F54DE988DA0ED905A9 -tenantId 9cfc42cb-51da-4055-87e9-b20a170b6ba3

.EXAMPLE
    # Process all sites and convert Organization links to direct permissions
    . .\Get-and-Remove-SPOSharingLinks.ps1 -mode Remediation -tenantName m365cpi13246019 -appId abc64618-283f-47ba-a185-50d935d51d57 -thumbprint B696FDCFE1453F3FBC6031F54DE988DA0ED905A9 -tenantId 9cfc42cb-51da-4055-87e9-b20a170b6ba3

.EXAMPLE
    # Process all sites and convert Organization links to direct permissions, and remove Anyone links
    . .\Get-and-Remove-SPOSharingLinks.ps1 -mode Remediation -removeAnyoneLinks $true -tenantName m365cpi13246019 -appId abc64618-283f-47ba-a185-50d935d51d57 -thumbprint B696FDCFE1453F3FBC6031F54DE988DA0ED905A9 -tenantId 9cfc42cb-51da-4055-87e9-b20a170b6ba3
#>

# ----------------------------------------------
# Set Variables
# ----------------------------------------------
param (
    # Tenant Name for your tenant
    [Parameter(Mandatory = $true)]
    [string]$tenantName,

    # Entra App ID for authentication
    [Parameter(Mandatory = $true)]
    [string]$appId,

    # Certificate thumbprint for authentication
    [Parameter(Mandatory = $true)]
    [string]$thumbprint,

    # Tenant ID for your tenant
    [Parameter(Mandatory = $true)]
    [string]$tenantId,

    # Path to the input file containing site URLs to scan
    [Parameter(Mandatory = $true)]
    [string]$inputFile,

    # Set to "Detection" for report mode, "Remediation" to convert Organization sharing links to direct permissions
    [ValidateSet('Detection', 'Remediation')]
    [string]$Mode = 'Detection',

    # If set to 'true', then Flexible Link Groups are not expanded and shown in the output.
    # Reduces discovery time, but direct sharing links are not presented
    [bool]$ignoreFlexibleLinkGroups = $true,

    # If set to 'true' then Anyone/Anonymous sharing links will be removed.
    # Only works if $Mode = 'Remediation'
    [bool]$removeAnyoneLinks = $false,

    # Path to save the output log file
    # If not specified then it will saved as 'SPO_SharingLinks_yyyyMMdd_HHmmss.txt'
    $logFilePath,

    # Path to save the output CSV file
    # If not specified then it will saved as 'SPO_SharingLinks_yyyyMMdd_HHmmss.txt'
    $outputFilePath,

    # Set to $true for verbose logging, $false for essential logging only
    # Default is false
    [switch]$debugLogging
)
# ----------------------------------------------
# Initialize Parameters - Do not change
# ----------------------------------------------
$sites = @()
$log = $null

# ----------------------------------------------
# Input / Output and Log Files
# ----------------------------------------------
$startime = Get-Date -Format 'yyyyMMdd_HHmmss'
if ([String]::IsNullOrEmpty($logFilePath)) {
    $logFilePath = ".\SPO_SharingLinks_$($startime).log"
}

if ([String]::IsNullOrEmpty($outputFilePath)) {
    $outputFilePath = ".\SPO_SharingLinks_$($startime).csv"
}
New-Item $logFilePath -ItemType File -ErrorAction Stop
New-Item $outputFilePath -ItemType File -ErrorAction Stop

# ----------------------------------------------
# Logging Function
# ----------------------------------------------
function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $Level = 'INFO' # INFO, DEBUG, ERROR
    )

    # Always log INFO and ERROR messages
    # Only log DEBUG messages when debug logging is enabled
    if ($Level -eq 'ERROR' -or $Level -eq 'INFO' -or ($Level -eq 'DEBUG' -and $debugLogging)) {
        if ($null -ne $LogName) {
            # log the date and time in the text file along with the data passed
            "$([DateTime]::Now.ToShortDateString()) $([string]([DateTime]::Now.TimeOfDay)) [$Level] : $LogEntryText" | Out-File -FilePath $LogName -Append
        }
    }
}

# ----------------------------------------------
# Logging Helper Functions
# ----------------------------------------------
function Write-InfoLog {
    param(
        [string] $LogName,
        [string] $LogEntryText
    )
    # Always log INFO messages
    Write-LogEntry -LogName $LogName -LogEntryText $LogEntryText -Level 'INFO'
}

function Write-DebugLog {
    param(
        [string] $LogName,
        [string] $LogEntryText
    )
    # Only log DEBUG messages when debug logging is enabled
    Write-LogEntry -LogName $LogName -LogEntryText $LogEntryText -Level 'DEBUG'
}

function Write-ErrorLog {
    param(
        [string] $LogName,
        [string] $LogEntryText
    )
    # Always log ERROR messages
    Write-LogEntry -LogName $LogName -LogEntryText $LogEntryText -Level 'ERROR'
}

# ----------------------------------------------
# PnP Version Detection and Graph Token Helper Function
# ----------------------------------------------
function Get-PnPGraphTokenCompatible {
    <#
    .SYNOPSIS
    Gets a Graph access token using the appropriate command based on PnP PowerShell version.
    Performs the check once and then re-uses the value

    .DESCRIPTION
    Automatically detects PnP PowerShell version and uses:
    - Get-PnPAccessToken for PnP PowerShell 3.0+
    - Get-PnPGraphAccessToken for PnP PowerShell 2.x and earlier
    #>

    if ($Script:PnpGraphTokenCompatible) {
        return & $Script:PnpGraphTokenCompatible
    }

    Write-DebugLog -LogName $Log -LogEntryText 'PnpGraphTokenCompatible not defined.'
    try {
        # Get the PnP PowerShell module version
        $pnpModule = Get-Module -Name 'PnP.PowerShell' -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1

        if (-not $pnpModule) {
            throw 'PnP.PowerShell module not found'
        }

        $majorVersion = $pnpModule.Version.Major
        Write-DebugLog -LogName $Log -LogEntryText "Detected PnP.PowerShell version: $($pnpModule.Version) (Major: $majorVersion)"

        if ($majorVersion -ge 3) {
            # PnP PowerShell 3.0+ uses Get-PnPAccessToken
            Write-DebugLog -LogName $Log -LogEntryText 'Using Get-PnPAccessToken for PnP PowerShell 3.0+'
            $Script:PnpGraphTokenCompatible = [scriptblock]::Create({ Get-PnPAccessToken })
        }
        else {
            # PnP PowerShell 2.x and earlier uses Get-PnPGraphAccessToken
            Write-DebugLog -LogName $Log -LogEntryText 'Using Get-PnPGraphAccessToken for PnP PowerShell 2.x'
            $Script:PnpGraphTokenCompatible = [scriptblock]::Create({ Get-PnPGraphAccessToken })
        }
        return & $Script:PnpGraphTokenCompatible
    }
    catch {
        # Fallback: try the newer command first, then the older one
        Write-DebugLog -LogName $Log -LogEntryText "Version detection failed, trying fallback approach: $_"

        try {
            Write-DebugLog -LogName $Log -LogEntryText 'Fallback: Attempting Get-PnPAccessToken (PnP 3.0+)'
            $Script:PnpGraphTokenCompatible = [scriptblock]::Create({ Get-PnPAccessToken })
        }
        catch {
            Write-DebugLog -LogName $Log -LogEntryText 'Fallback: Attempting Get-PnPGraphAccessToken (PnP 2.x)'
            $Script:PnpGraphTokenCompatible = [scriptblock]::Create({ Get-PnPGraphAccessToken })
        }
        return & $Script:PnpGraphTokenCompatible
    }
}

# ----------------------------------------------
# Determine Script Operation Mode and Auto-Configure Settings
# ----------------------------------------------

# Validate and normalize the Mode parameter
if ($Mode -notin @('Detection', 'Remediation')) {
    Write-Host "Invalid Mode specified: '$Mode'. Must be 'Detection' or 'Remediation'. Defaulting to 'Detection'." -ForegroundColor Red
    $Mode = 'Detection'
}

# Set internal variables based on Mode
$convertOrganizationLinks = ($Mode -eq 'Remediation')
$RemoveSharingLink = ($Mode -eq 'Remediation')  # Always remove sharing links in Remediation mode
$removeAnyoneLinks = ($removeAnyoneLinks -and $Mode -eq 'Remediation')
$scriptMode = $Mode.ToUpper()

# Auto-enable cleanup when in remediation mode
if ($convertOrganizationLinks) {
    $cleanupCorruptedSharingGroups = $true
    Write-InfoLog -LogName $Log -LogEntryText 'Auto-enabled cleanup of corrupted sharing groups because remediation mode is active'
}
else {
    $cleanupCorruptedSharingGroups = $false
}

Write-Host "Script is running in $scriptMode mode" -ForegroundColor $(if ($convertOrganizationLinks) {
        'Yellow'
    }
    else {
        'Cyan'
    })
Write-InfoLog -LogName $Log -LogEntryText "Script is running in $scriptMode mode - $(if ($convertOrganizationLinks) { 'Converting Organization links to direct permissions and removing Organization sharing groups (flexible links are preserved)' } else { 'Only detecting and inventorying sharing links, no modifications will be made' })"


# ----------------------------------------------
# Connection Parameters
# ----------------------------------------------
$connectionParams = @{
    ClientId      = $appID
    Thumbprint    = $thumbprint
    Tenant        = $tenantId
    WarningAction = 'SilentlyContinue'
}

# ----------------------------------------------
# Throttling Handling Function
# ----------------------------------------------
function Invoke-WithThrottleHandling {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock] $ScriptBlock,

        [Parameter(Mandatory = $false)]
        [int] $MaxRetries = 5,

        [Parameter(Mandatory = $false)]
        [string] $Operation = 'SharePoint Operation'
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
                    $retryAfterHeader = $_.Exception.Response.Headers['Retry-After']

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
            elseif ($errorMessage -match 'throttl|Too many requests|429|503|Request limit exceeded') {
                $isThrottling = $true

                # Extract wait time from error message if available
                if ($errorMessage -match 'Try again in (\d+) (seconds|minutes)') {
                    $timeValue = [int]$matches[1]
                    $timeUnit = $matches[2]

                    $waitTime = if ($timeUnit -eq 'minutes') {
                        $timeValue * 60
                    }
                    else {
                        $timeValue
                    }
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
if ($inputFile -and (Test-Path -Path $inputFile)) {
    Write-Host "Processing input file: $inputFile" -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "Processing input file: $inputFile"

    try {
        # Check if this is the script's CSV output format or a simple URL list
        $firstLine = Get-Content -Path $inputFile -TotalCount 1

        if ($firstLine -and $firstLine.Contains('Sharing Group Name')) {
            # This is the script's CSV output format
            Write-Host "Detected script's CSV output format - will process Organization sharing links only" -ForegroundColor Cyan
            Write-InfoLog -LogName $Log -LogEntryText "Input file detected as script's CSV output format"

            # Import the full CSV and filter for Organization sharing links only
            $csvData = Import-Csv -Path $inputFile
            $organizationEntries = $csvData | Where-Object {
                $_.'Sharing Group Name' -like '*Organization*' -and
                -not [string]::IsNullOrWhiteSpace($_.'Site URL')
            }

            if ($organizationEntries.Count -eq 0) {
                Write-Host 'No Organization sharing links found in the input CSV file' -ForegroundColor Yellow
                Write-InfoLog -LogName $Log -LogEntryText 'No Organization sharing links found in input CSV'
                $sites = @()
            }
            else {
                # Group by Site URL to get unique sites and force conversion mode
                $siteGroups = $organizationEntries | Group-Object 'Site URL'
                $sites = $siteGroups | ForEach-Object { [PSCustomObject]@{ URL = $_.Name } }

                # Auto-enable remediation when using CSV output
                if ($Mode -eq 'Detection') {
                    Write-Host 'Auto-enabling Remediation mode for CSV input containing Organization links' -ForegroundColor Green
                    $Mode = 'Remediation'
                    $convertOrganizationLinks = $true
                    $RemoveSharingLink = $true
                    $cleanupCorruptedSharingGroups = $true

                    # Update the script mode to reflect the change
                    $scriptMode = 'REMEDIATION'
                    Write-Host "Updated script mode to $scriptMode" -ForegroundColor Yellow
                    Write-InfoLog -LogName $Log -LogEntryText 'Auto-enabled remediation mode for CSV input containing Organization links'
                }

                Write-Host "Found $($sites.Count) sites with Organization sharing links for remediation" -ForegroundColor Green
                Write-InfoLog -LogName $Log -LogEntryText "Parsed $($sites.Count) sites with Organization sharing links from CSV input"
            }
        }
        else {
            # This is a simple site URL list
            Write-Host 'Input file appears to be a simple site URL list' -ForegroundColor Yellow
            Write-InfoLog -LogName $Log -LogEntryText 'Input file detected as simple site URL list'
            $sites = Import-Csv -Path $inputFile -Header 'URL'
        }
    }
    catch {
        Write-Host "Error reading input file '$inputFile': $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error reading input file '$inputFile': $_"
        exit
    }
}
else {
    Write-Host 'Getting site list from tenant (this might take a while)...' -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText 'Getting sites using Get-PnPTenantSite (no input file specified or found)'
    try {
        # Get sites with optimized filtering to reduce memory usage and improve performance
        $sites = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPTenantSite -IncludeOneDriveSites:$false | Where-Object {
                $_.Template -notmatch 'SRCHCEN|MYSITE|APPCATALOG|PWS|POINTPUBLISHINGTOPIC|SPSMSITEHOST|EHS|REVIEWCTR|TENANTADMIN' -and
                $_.Status -eq 'Active' -and
                -not [string]::IsNullOrEmpty($_.Url)
            }
        } -Operation 'Get-PnPTenantSite with optimized filtering'

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
$sharingLinksHeaders = 'Site URL,Site Owner,IB Mode,IB Segment,Site Template,Sharing Group Name,Sharing Link Members,File URL,File Owner,Filename,SharingType,Sharing Link URL,Link Expiration Date,IsTeamsConnected,SharingCapability,Last Content Modified,Search Status,Link Removed'
Set-Content -Path $sharingLinksOutputFile -Value $sharingLinksHeaders
Write-InfoLog -LogName $Log -LogEntryText "Initialized sharing links output file: $sharingLinksOutputFile"

# ----------------------------------------------
# Function to handle consolidated site data
# ----------------------------------------------
function Update-SiteCollectionData {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [object] $SiteProperties,
        [string] $SPGroupName = '',
        # --- Parameters for SP User ---
        [string] $AssociatedSPGroup = '',
        [string] $SPUserName = '',
        [string] $SPUserTitle = '',
        [string] $SPUserEmail = ''
    )

    # Create site entry if it doesn't exist
    if (-not $siteCollectionData.ContainsKey($SiteUrl)) {
        $siteCollectionData[$SiteUrl] = @{
            'URL'                     = $SiteProperties.Url
            'Owner'                   = $SiteProperties.Owner
            'IB Mode'                 = ($SiteProperties.InformationBarrierMode -join ',')
            'IB Segment'              = ($SiteProperties.InformationBarrierSegments -join ',')
            'Template'                = $SiteProperties.Template
            'SharingCapability'       = $SiteProperties.SharingCapability
            'IsTeamsConnected'        = $SiteProperties.IsTeamsConnected
            'LastContentModifiedDate' = $SiteProperties.LastContentModifiedDate
            # Site-specific lists
            'SP Groups On Site'       = [System.Collections.Generic.List[string]]::new()
            'SP Users'                = [System.Collections.Generic.List[PSObject]]::new()
            'Has Sharing Links'       = $false # Property to track if sharing links are being used
            'Link Removal Status'     = @{} # Track which sharing groups had their links removed
        }
    }

    # Check for SharingLinks groups
    if (-not [string]::IsNullOrWhiteSpace($SPGroupName) -and $SPGroupName -like 'SharingLinks*') {
        $siteCollectionData[$SiteUrl]['Has Sharing Links'] = $true

        # Initialize link removal status to False for all sharing groups by default
        if (-not $siteCollectionData[$SiteUrl]['Link Removal Status'].ContainsKey($SPGroupName)) {
            $siteCollectionData[$SiteUrl]['Link Removal Status'][$SPGroupName] = $false
        }
    }

    # Add SP Group if provided and not already present for this site
    if (-not [string]::IsNullOrWhiteSpace($SPGroupName)) {
        if (-not $siteCollectionData[$SiteUrl]['SP Groups On Site'].Contains($SPGroupName)) {
            $siteCollectionData[$SiteUrl]['SP Groups On Site'].Add($SPGroupName)
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
        $siteCollectionData[$SiteUrl]['SP Users'].Add($userObject)

        # Debug: Log when we store sharing group members
        if ($AssociatedSPGroup -like 'SharingLinks*') {
            Write-DebugLog -LogName $Log -LogEntryText "STORED user in site data: Group='$AssociatedSPGroup', Name='$SPUserName', Title='$SPUserTitle', Email='$SPUserEmail'"
        }
    }
    else {
        # Debug: Log when we skip storing a user
        if ($AssociatedSPGroup -like 'SharingLinks*') {
            Write-DebugLog -LogName $Log -LogEntryText "SKIPPED storing user: SPUserName empty: $([string]::IsNullOrWhiteSpace($SPUserName)), AssociatedSPGroup empty: $([string]::IsNullOrWhiteSpace($AssociatedSPGroup)), Values: Name='$SPUserName', Group='$AssociatedSPGroup'"
        }
    }
}

# ----------------------------------------------
# Function to update link removal status for a sharing group
# ----------------------------------------------
function Update-LinkRemovalStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [Parameter(Mandatory = $true)]
        [string] $SharingGroupName,
        [Parameter(Mandatory = $true)]
        [bool] $WasRemoved
    )

    if ($siteCollectionData.ContainsKey($SiteUrl)) {
        $siteCollectionData[$SiteUrl]['Link Removal Status'][$SharingGroupName] = $WasRemoved
        Write-DebugLog -LogName $Log -LogEntryText "Updated link removal status for $SharingGroupName on $SiteUrl : $WasRemoved"
    }
}

# ----------------------------------------------
# Function to process and write sharing links for a site
# ----------------------------------------------
function Write-SiteSharingLinks {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        [object] $SiteData
    )

    # Check if this site has sharing links groups
    $sharingLinkGroups = $SiteData.'SP Groups On Site' | Where-Object { $_ -like 'SharingLinks*' }

    if ($sharingLinkGroups.Count -gt 0) {
        Write-Host "  Processing $($sharingLinkGroups.Count) sharing link groups for site: $SiteUrl" -ForegroundColor Yellow
        Write-InfoLog -LogName $Log -LogEntryText "Processing $($sharingLinkGroups.Count) sharing link groups for site: $SiteUrl"

        foreach ($sharingGroup in $sharingLinkGroups) {
            # Get users in this sharing links group
            $groupMembers = $SiteData.'SP Users' | Where-Object { $_.AssociatedSPGroup -eq $sharingGroup }

            # Debug: Log what members we found for this sharing group
            Write-DebugLog -LogName $Log -LogEntryText "Processing sharing group '$sharingGroup' - found $($groupMembers.Count) members in site data"

            # Debug: Also show ALL users for this site to verify data storage
            $allSiteUsers = $SiteData.'SP Users'
            $allSharingUsers = $allSiteUsers | Where-Object { $_.AssociatedSPGroup -like 'SharingLinks*' }
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
                    $emailStr = if ($_.Email) {
                        $_.Email | Out-String -NoNewline
                    }
                    else {
                        ''
                    }
                    "$($_.Name) <$emailStr>"
                }) -join ';'
            }
            else {
                # Check if the file was not searchable to provide better context for empty member lists
                if ($searchStatus -eq 'Not Found in Search') {
                    'Not Searchable'
                }
                elseif ($searchStatus -eq 'Search Error') {
                    'Search Error'
                }
                else {
                    'No members'
                }
            }

            # Get document details if available
            $documentUrl = 'Not found'
            $documentOwner = 'Not found'
            $documentItemType = 'Not found'
            $sharingLinkUrl = 'Not found'
            $linkExpirationDate = 'Not found'
            $searchStatus = 'Not Searched'
            if ($SiteData.ContainsKey('DocumentDetails') -and $SiteData['DocumentDetails'].ContainsKey($sharingGroup)) {
                $documentUrl = $SiteData['DocumentDetails'][$sharingGroup]['DocumentUrl']
                $documentOwner = $SiteData['DocumentDetails'][$sharingGroup]['DocumentOwner']
                $documentItemType = $SiteData['DocumentDetails'][$sharingGroup]['DocumentItemType']
                $documentLastModified = $SiteData['DocumentDetails'][$sharingGroup]['DocumentLastModified']
                $sharingLinkUrl = $SiteData['DocumentDetails'][$sharingGroup]['SharingLinkUrl']
                $linkExpirationDate = $SiteData['DocumentDetails'][$sharingGroup]['ExpirationDate']
                $searchStatus = $SiteData['DocumentDetails'][$sharingGroup]['SearchStatus']
                Write-DebugLog -LogName $Log -LogEntryText "Retrieved document details for $sharingGroup - URL: $documentUrl, Owner: $documentOwner, Type: $documentItemType, LinkURL: $sharingLinkUrl, Expiration: $linkExpirationDate, SearchStatus: $searchStatus"
            }
            else {
                Write-DebugLog -LogName $Log -LogEntryText "No document details found for sharing group: $sharingGroup. DocumentDetails exists: $($SiteData.ContainsKey('DocumentDetails')), Group key exists: $(if ($SiteData.ContainsKey('DocumentDetails')) { $SiteData['DocumentDetails'].ContainsKey($sharingGroup) } else { 'N/A' })"
            }

            # Get link removal status
            $linkRemoved = 'False'
            if ($SiteData.ContainsKey('Link Removal Status') -and $SiteData['Link Removal Status'].ContainsKey($sharingGroup)) {
                $linkRemoved = if ($SiteData['Link Removal Status'][$sharingGroup]) {
                    'True'
                }
                else {
                    'False'
                }
            }

            # Extract filename from the document URL
            $filename = 'Not found'
            if ($documentUrl -ne 'Not found' -and $documentUrl -ne 'Not Searchable' -and $documentUrl -ne 'Search Error' -and -not [string]::IsNullOrWhiteSpace($documentUrl)) {
                try {
                    if ($documentUrl -match 'DispForm\.aspx\?ID=(\d+)') {
                        # This is a list item - try to get a meaningful name
                        # For list items, we'll use "List Item" + ID as the filename
                        $itemId = $matches[1]
                        $filename = "List Item $itemId"

                        # Try to extract list name for better context
                        if ($documentUrl -match '/Lists/([^/]+)/DispForm\.aspx') {
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
                    $filename = 'Extraction Error'
                }
            }
            elseif ($documentUrl -eq 'Not Searchable') {
                $filename = 'Not Searchable'
            }
            elseif ($documentUrl -eq 'Search Error') {
                $filename = 'Search Error'
            }

            # Determine sharing type based on sharing group name
            $sharingType = 'Unknown'
            if ($sharingGroup -like '*Flexible*') {
                $sharingType = 'Flexible'
            }
            elseif ($sharingGroup -like '*Organization*') {
                $sharingType = 'Organization'
            }
            elseif ($sharingGroup -like '*Anonymous*') {
                $sharingType = 'Anonymous'
            }

            # Create CSV line
            $csvLine = [PSCustomObject]@{
                'Site URL'              = $SiteData.URL
                'Site Owner'            = $SiteData.Owner
                'IB Mode'               = $SiteData.'IB Mode'
                'IB Segment'            = $SiteData.'IB Segment'
                'Site Template'         = $SiteData.Template
                'Sharing Group Name'    = $sharingGroup
                'Sharing Link Members'  = $membersFormatted
                'File URL'              = $documentUrl
                'File Owner'            = $documentOwner
                'Filename'              = $filename
                'SharingType'           = $sharingType
                'Sharing Link URL'      = $sharingLinkUrl
                'Link Expiration Date'  = $linkExpirationDate
                'IsTeamsConnected'      = $SiteData.IsTeamsConnected
                'SharingCapability'     = $SiteData.SharingCapability
                'Last Content Modified' = $documentLastModified #$SiteData.LastContentModifiedDate
                'Search Status'         = $searchStatus
                'Link Removed'          = $linkRemoved
            }

            # Write directly to the CSV file
            $csvLine | Export-Csv -Path $sharingLinksOutputFile -Append -NoTypeInformation -Force
            Write-DebugLog -LogName $Log -LogEntryText "  Wrote sharing link data for group: $sharingGroup"
        }
    }
}

# ----------------------------------------------
# Function to convert Organization sharing links to direct permissions
# ----------------------------------------------
function Convert-OrganizationSharingLinks {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl
    )

    Write-Host "  Checking for Organization sharing links on site: $SiteUrl" -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "Checking for Organization sharing links on site: $SiteUrl"

    try {
        # Connect to the specific site
        $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if (-not $currentConnection -or $currentConnection.Url -ne $SiteUrl) {
            Connect-PnPOnline -Url $SiteUrl @connectionParams -ErrorAction Stop
        }

        # Get all SharePoint groups that contain "Organization" in the name
        $organizationGroups = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPGroup | Where-Object { $_.Title -like '*Organization*' }
        } -Operation "Get Organization groups for $SiteUrl"

        if ($organizationGroups.Count -eq 0) {
            Write-DebugLog -LogName $Log -LogEntryText "No Organization sharing groups found on site: $SiteUrl"
            return
        }

        Write-Host "    Found $($organizationGroups.Count) Organization sharing groups" -ForegroundColor Green
        Write-InfoLog -LogName $Log -LogEntryText "Found $($organizationGroups.Count) Organization sharing groups on site: $SiteUrl"

        # Get a list of the site and subsite URLs.
        # This is used to find the link in the appropriate site to increase the chance of a successful removal using PnP Methods
        # As Get-PnpFile only works if you are connected to the right site.
        Write-DebugLog -LogName $Log -LogEntryText 'Retrieving Site and Subsite URLs'
        $subsites = Get-PnPSubWeb -Recurse | Select-Object Title, Url
        $subsites += Get-PnPWeb -Includes Title, Url | Select-Object Title, Url

        foreach ($orgGroup in $organizationGroups) {
            $groupName = $orgGroup.Title
            Write-Host "    Processing Organization group: $groupName" -ForegroundColor Cyan
            Write-DebugLog -LogName $Log -LogEntryText "Processing Organization group: $groupName"

            # Determine permission level based on group name
            $permissionLevel = ''
            if ($groupName -like '*OrganizationEdit*') {
                $permissionLevel = 'Edit'
            }
            elseif ($groupName -like '*OrganizationView*') {
                $permissionLevel = 'Read'
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
            $documentUrl = ''
            $documentId = ''

            if ($groupName -match 'SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.') {
                $documentId = $matches[1]
                Write-DebugLog -LogName $Log -LogEntryText "Extracted document ID: $documentId from group: $groupName"

                # Try to find the document using existing site collection data first
                if ($siteCollectionData[$SiteUrl].ContainsKey('DocumentDetails') -and
                    $siteCollectionData[$SiteUrl]['DocumentDetails'].ContainsKey($groupName) -and
                    -not [string]::IsNullOrWhiteSpace($siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['DocumentUrl'])) {
                    $documentUrl = $siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['DocumentUrl']
                    Write-DebugLog -LogName $Log -LogEntryText "Found document URL: $documentUrl"
                }
                else {
                    Write-DebugLog -LogName $Log -LogEntryText "No document URL found in site collection data for group: $groupName"
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

            # Process each member using the WORKING approach from GitHub script
            foreach ($member in $groupMembers) {
                if (!$member -or !$member.LoginName) {
                    continue
                }

                try {
                    Write-Host "        Processing member: $($member.Title)" -ForegroundColor White
                    Write-DebugLog -LogName $Log -LogEntryText "Processing member: $($member.Title) ($($member.LoginName))"

                    # STEP 1: Remove user from the sharing group FIRST (key difference from the broken version)
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

                    # STEP 2: Grant direct permissions to the document if we found it
                    if (-not [string]::IsNullOrWhiteSpace($documentUrl)) {
                        try {
                            # Parse the document URL to get the relative URL
                            $uri = [System.Uri]$documentUrl
                            $relativePath = $uri.AbsolutePath

                            # Validate that we have a non-empty relative path
                            if ([string]::IsNullOrWhiteSpace($relativePath)) {
                                Write-DebugLog -LogName $Log -LogEntryText "Warning: Empty relative path from document URL: $documentUrl"
                                throw 'Invalid document URL - empty relative path'
                            }

                            Write-DebugLog -LogName $Log -LogEntryText "Attempting to grant permissions for document URL: $documentUrl"
                            Write-DebugLog -LogName $Log -LogEntryText "Parsed relative path: $relativePath"

                            # Check if this is a SharePoint list item (with DispForm.aspx) vs a document library file
                            # Note: Need to check the full URL, not just the path, for DispForm.aspx pattern
                            if ($documentUrl -match 'DispForm\.aspx\?ID=(\d+)' -or $relativePath -match 'DispForm\.aspx\?ID=(\d+)') {
                                # This is a SharePoint list item - handle it differently
                                $itemId = $matches[1]
                                Write-DebugLog -LogName $Log -LogEntryText "Detected SharePoint list item with ID: $itemId from URL: $documentUrl"

                                # Extract list name from the URL path - improved logic
                                $listName = ''

                                # Try different patterns to extract the list name
                                if ($documentUrl -match '/Lists/([^/]+)/DispForm\.aspx') {
                                    $listName = $matches[1]
                                    Write-DebugLog -LogName $Log -LogEntryText "Extracted list name from /Lists/ pattern: $listName"
                                }
                                elseif ($documentUrl -match '/sites/[^/]+/([^/]+)/DispForm\.aspx') {
                                    $listName = $matches[1]
                                    Write-DebugLog -LogName $Log -LogEntryText "Extracted list name from site pattern: $listName"
                                }
                                else {
                                    # Fallback: parse the path manually
                                    $pathParts = $relativePath.Split('/')

                                    # Find the part before 'DispForm.aspx'
                                    for ($i = 0; $i -lt $pathParts.Length; $i++) {
                                        if ($pathParts[$i] -eq 'DispForm.aspx') {
                                            if ($i -gt 0) {
                                                $listName = $pathParts[$i - 1]
                                                Write-DebugLog -LogName $Log -LogEntryText "Extracted list name from path parsing: $listName"
                                            }
                                            break
                                        }
                                    }
                                }

                                if (-not [string]::IsNullOrWhiteSpace($listName)) {
                                    Write-DebugLog -LogName $Log -LogEntryText "Attempting to grant permissions to list item - List: '$listName', Item ID: $itemId"

                                    try {
                                        # Use PnP to grant permissions directly to the list item
                                        Invoke-WithThrottleHandling -ScriptBlock {
                                            Set-PnPListItemPermission -List $listName -Identity $itemId -User $member.LoginName -AddRole $permissionLevel
                                        } -Operation "Grant $permissionLevel permission to $($member.LoginName) for list item"

                                        Write-Host "          Granted direct $permissionLevel permission to list item (List: $listName, ID: $itemId)" -ForegroundColor Green
                                        Write-InfoLog -LogName $Log -LogEntryText "Granted direct $permissionLevel permission to $($member.LoginName) for list item: $documentUrl"
                                    }
                                    catch {
                                        Write-DebugLog -LogName $Log -LogEntryText "Direct list item permission failed for list '$listName', item ID '$itemId': $($_.Exception.Message)"
                                        throw "Could not grant permissions to list item '$listName' (ID: $itemId): $($_.Exception.Message)"
                                    }
                                }
                                else {
                                    $errorMsg = "Could not extract list name from SharePoint list item URL: $documentUrl (relative path: $relativePath)"
                                    Write-DebugLog -LogName $Log -LogEntryText $errorMsg
                                    throw $errorMsg
                                }
                            }
                            else {
                                # This is a document library file - try to grant permissions to the file
                                Write-DebugLog -LogName $Log -LogEntryText 'Processing as document library file'

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

                                    # Fallback: Grant permissions at site level
                                    throw 'CSOM failed, falling back to site-level permissions'
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

            # STEP 3: Always remove empty Organization sharing groups (critical for list items)
            Write-Host "      Checking if Organization sharing group is empty: $groupName" -ForegroundColor Yellow
            Write-DebugLog -LogName $Log -LogEntryText "Checking if Organization sharing group is empty after member removal: $groupName"

            try {
                # Check if the group is empty after removing members
                $remainingMembers = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPGroupMember -Identity $orgGroup.Id -ErrorAction SilentlyContinue
                } -Operation "Check remaining members in Organization group $groupName"

                if ($remainingMembers -and $remainingMembers.Count -gt 0) {
                    Write-LogEntry -LogName $Log -LogEntryText "Warning: Group $groupName still has $($remainingMembers.Count) members, will not remove" -Level 'INFO'
                    Write-Host '      Warning: Group still has members, skipping removal' -ForegroundColor Yellow
                }
                else {
                    # Group is empty - remove it to prevent corruption (especially important for list items)
                    Write-Host "      Removing empty Organization sharing group: $groupName" -ForegroundColor Green
                    Write-InfoLog -LogName $Log -LogEntryText "Attempting to remove empty Organization sharing group: $groupName"

                    try {
                        Invoke-WithThrottleHandling -ScriptBlock {
                            # First check if group still exists
                            $groupCheck = Get-PnPGroup -Identity $orgGroup.Id -ErrorAction SilentlyContinue
                            if ($groupCheck) {
                                Remove-PnPGroup -Identity $orgGroup.Id -Force
                                Write-Host "        Successfully removed empty Organization sharing group: $groupName" -ForegroundColor Green
                                Write-InfoLog -LogName $Log -LogEntryText "Successfully removed empty Organization sharing group: $groupName"
                            }
                            else {
                                Write-LogEntry -LogName $Log -LogEntryText "Group $groupName no longer exists, may have already been removed" -Level 'INFO'
                                Write-Host '        Group no longer exists (may have already been removed)' -ForegroundColor Yellow
                            }
                        } -Operation "Remove empty Organization sharing group $groupName"

                        Update-LinkRemovalStatus -SiteUrl $SiteUrl -SharingGroupName $groupName -WasRemoved $true
                    }
                    catch {
                        Write-Host "        Warning: Could not remove sharing group $groupName : $_" -ForegroundColor Red
                        Write-ErrorLog -LogName $Log -LogEntryText "Failed to remove sharing group $groupName : $_"
                        Update-LinkRemovalStatus -SiteUrl $SiteUrl -SharingGroupName $groupName -WasRemoved $false
                    }
                }
            }
            catch {
                Write-Host "      Warning: Error checking/removing Organization sharing group $groupName : $_" -ForegroundColor Red
                Write-ErrorLog -LogName $Log -LogEntryText "Error checking/removing Organization sharing group $groupName : $_"
                Update-LinkRemovalStatus -SiteUrl $SiteUrl -SharingGroupName $groupName -WasRemoved $false
            }

            # STEP 4: Optional sharing link removal (for documents only)
            if ($RemoveSharingLink) {
                # First verify the group is empty before removing
                $remainingMembers = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPGroupMember -Identity $orgGroup.Id -ErrorAction SilentlyContinue
                } -Operation "Check remaining members in Organization group $groupName"

                if ($remainingMembers -and $remainingMembers.Count -gt 0) {
                    Write-LogEntry -LogName $Log -LogEntryText "Warning: Group $groupName still has $($remainingMembers.Count) members, will not remove" -Level 'INFO'
                    Write-Host '      Warning: Group still has members, skipping removal' -ForegroundColor Yellow
                }
                else {
                    Remove-SharingLink -siteUrl $siteUrl -groupName $orgGroup.Title -groupId $orgGroup.Id -documentId $documentId -documentUrl $documentUrl
                }
            }
            else {
                Write-Host "      Preserving Organization sharing group: $groupName (RemoveSharingLink is disabled)" -ForegroundColor Cyan
                Write-InfoLog -LogName $Log -LogEntryText "Preserving Organization sharing group: $groupName because RemoveSharingLink is disabled"
                Update-LinkRemovalStatus -SiteUrl $SiteUrl -SharingGroupName $groupName -WasRemoved $false
            }
        }
    }
    catch {
        Write-Host "  Error processing Organization sharing links for site $SiteUrl : $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error processing Organization sharing links for site $SiteUrl : $_"
    }
}

function Remove-CorruptedSharingGroups {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl
    )

    Write-Host "  Checking for corrupted sharing groups on site: $SiteUrl (excluding flexible links)" -ForegroundColor Yellow
    Write-LogEntry -LogName $Log -LogEntryText "Checking for corrupted sharing groups on site: $SiteUrl (excluding flexible sharing links)" -Level 'INFO'

    try {
        # Connect to the specific site
        $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if (-not $currentConnection -or $currentConnection.Url -ne $SiteUrl) {
            Connect-PnPOnline -Url $SiteUrl @connectionParams -ErrorAction Stop
        }

        # Get all SharePoint groups that look like sharing groups
        $allSharingGroups = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPGroup | Where-Object { $_.Title -like 'SharingLinks*' }
        } -Operation "Get all sharing groups for $SiteUrl"

        if ($allSharingGroups.Count -eq 0) {
            Write-LogEntry -LogName $Log -LogEntryText "No sharing groups found on site: $SiteUrl" -Level 'DEBUG'
            return
        }

        $corruptedGroupsRemoved = 0

        foreach ($sharingGroup in $allSharingGroups) {
            try {
                # Skip flexible sharing links - only clean up Organization sharing groups in remediation mode
                if ($sharingGroup.Title -like '*Flexible*') {
                    Write-DebugLog -LogName $Log -LogEntryText "Skipping flexible sharing group from cleanup: $($sharingGroup.Title)"
                    continue
                }

                # Check if group has any members
                $groupMembers = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPGroupMember -Identity $sharingGroup.Id -ErrorAction SilentlyContinue
                } -Operation "Check members in sharing group $($sharingGroup.Title)"

                # If group has no members, it's likely corrupted
                if (-not $groupMembers -or $groupMembers.Count -eq 0) {
                    Write-Host "    Found empty sharing group: $($sharingGroup.Title)" -ForegroundColor Yellow
                    Write-LogEntry -LogName $Log -LogEntryText "Found potentially corrupted empty sharing group: $($sharingGroup.Title)" -Level 'INFO'

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
# Function to remove Anonymous sharing links
# ----------------------------------------------
function Remove-AnyoneSharingLinks {
    param(
        [Parameter(Mandatory = $true)]
        [string] $siteUrl
    )

    Write-Host "  Checking for Anyone sharing links on site: $SiteUrl" -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText "Checking for Anyone sharing links on site: $SiteUrl"

    try {
        # Connect to the specific site
        $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if (-not $currentConnection -or $currentConnection.Url -ne $SiteUrl) {
            Connect-PnPOnline -Url $SiteUrl @connectionParams -ErrorAction Stop
        }

        # Get all SharePoint groups that contain "Anonymous" in the name
        $anonymousGroups = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPGroup | Where-Object { $_.Title -like '*Anonymous*' }
        } -Operation "Get Anonymous groups for $SiteUrl"

        if ($anonymousGroups.Count -eq 0) {
            Write-DebugLog -LogName $Log -LogEntryText "No Anonymous sharing groups found on site: $SiteUrl"
            return
        }

        # Get a list of the site and subsite URLs.
        # This is used to find the link in the appropriate site to increase the chance of a successful removal using PnP Methods
        # As Get-PnpFile only works if you are connected to the right site.
        Write-DebugLog -LogName $Log -LogEntryText 'Retrieving Site and Subsite URLs'
        $subsites = Get-PnPSubWeb -Recurse | Select-Object Title, Url
        $subsites += Get-PnPWeb -Includes Title, Url | Select-Object Title, Url

        foreach ($anonGroup in $anonymousGroups) {
            $groupName = $anonGroup.Title
            Write-Host "    Processing Anonymous group: $groupName" -ForegroundColor Cyan
            Write-DebugLog -LogName $Log -LogEntryText "Processing Anonymous group: $groupName"

            # Extract document information from group name
            $documentUrl = ''
            $documentId = ''

            if ($groupName -match 'SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.Anonymous.*([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
                $documentId = $matches[1]
                Write-DebugLog -LogName $Log -LogEntryText "Extracted document ID: $documentId from group: $groupName"

                # Try to find the document using existing site collection data first
                if ($siteCollectionData[$SiteUrl].ContainsKey('DocumentDetails') -and
                    $siteCollectionData[$SiteUrl]['DocumentDetails'].ContainsKey($groupName) -and
                    -not [string]::IsNullOrWhiteSpace($siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['DocumentUrl'])) {
                    $documentUrl = $siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['DocumentUrl']
                    Write-DebugLog -LogName $Log -LogEntryText "Found document URL: $documentUrl"
                }
                else {
                    Write-DebugLog -LogName $Log -LogEntryText "No document URL found in site collection data for group: $groupName"
                }

            }
            else {
                Write-DebugLog -LogName $Log -LogEntryText "Could not extract document ID from group name: $groupName"
            }

            if ($RemoveSharingLink) {
                # Update-LinkRemovalStatus is called as part of the following function
                Remove-SharingLink -siteUrl $siteUrl -groupName $anonGroup.Title -groupId $anonGroup.Id -documentId $documentId -documentUrl $documentUrl -allSiteUrls $subsites.Url
            }
            else {
                Write-Host "      Preserving Anonymous sharing group: $groupName (RemoveSharingLink is disabled)" -ForegroundColor Cyan
                Write-InfoLog -LogName $Log -LogEntryText "Preserving Anonymous sharing group: $groupName because RemoveSharingLink is disabled"
                Update-LinkRemovalStatus -SiteUrl $SiteUrl -SharingGroupName $groupName -WasRemoved $false
            }

        }
    }
    catch {
        Write-Host "  Error processing Anonymous sharing links for site $SiteUrl : $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error processing Anonymous sharing links for site $SiteUrl : $_"
    }
}

# ----------------------------------------------
# Function to attempt the removal of a sharing link from an object
# ----------------------------------------------
function Remove-SharingLink {
    param (
        [string]$siteUrl,
        [string]$groupName,
        [string]$documentId,
        [string]$documentUrl,

        # Supply a list of all site and subsite URLs
        [string[]]$allSiteUrls
    )

    Write-Host "      Attempting optional sharing link removal: $groupName" -ForegroundColor Cyan
    Write-DebugLog -LogName $Log -LogEntryText "Attempting optional sharing link removal: $groupName"

    $sharingLinkRemoved = $false

    try {
        # Try to remove using UnshareLink if we have document details
        if ($documentUrl -and $documentId -and $documentUrl -ne 'Not Searchable') {
            try {
                Write-DebugLog -LogName $Log -LogEntryText 'Attempting to unshare link using PnP PowerShell methods'

                $result = Invoke-WithThrottleHandling -ScriptBlock {
                    # Parse document URL to get relative path
                    $uri = [System.Uri]$documentUrl
                    $relativePath = $uri.AbsolutePath

                    $linkRemoved = $false

                    # Check if this is a list item (DispForm.aspx) or a regular file
                    # Note: Need to check the full URL, not just the path, for DispForm.aspx pattern
                    if ($documentUrl -match 'DispForm\.aspx\?ID=(\d+)' -or $relativePath -match 'DispForm\.aspx\?ID=(\d+)') {
                        # This is a list item - we need to use the document ID directly
                        Write-DebugLog -LogName $Log -LogEntryText "Detected list item, using document ID $documentId for sharing link operations. Full URL: $documentUrl"

                        try {
                            # For list items, try to get sharing links using the document ID directly
                            Write-DebugLog -LogName $Log -LogEntryText "Attempting to get sharing links for list item with document ID: $documentId"

                            $sharingLinks = Invoke-WithThrottleHandling -ScriptBlock {
                                Get-PnPFileSharingLink -Identity $documentId -ErrorAction SilentlyContinue
                            } -Operation "Get sharing links for list item with document ID $documentId"

                            if ($sharingLinks -and $sharingLinks.Count -gt 0) {
                                Write-DebugLog -LogName $Log -LogEntryText "Found $($sharingLinks.Count) sharing links for list item with document ID: $documentId"

                                # Debug: Show all sharing link IDs
                                $linkIds = ($sharingLinks | ForEach-Object { $_.Id }) -join ', '
                                Write-DebugLog -LogName $Log -LogEntryText "Sharing link IDs found: $linkIds"
                                Write-DebugLog -LogName $Log -LogEntryText "Looking for sharing link matching group: $groupName"

                                foreach ($sharingLink in $sharingLinks) {
                                    Write-DebugLog -LogName $Log -LogEntryText "Checking sharing link ID: $($sharingLink.Id) against group: $groupName"

                                    # Try to match the sharing link with our group
                                    # The group name contains the document ID, so try multiple approaches
                                    $isMatch = $false

                                    if ($sharingLink.Id -and $groupName -like "*$($sharingLink.Id)*") {
                                        $isMatch = $true
                                        Write-DebugLog -LogName $Log -LogEntryText 'Match found using sharing link ID in group name'
                                    }
                                    elseif ($documentId -and $sharingLink.Id -eq $documentId) {
                                        $isMatch = $true
                                        Write-DebugLog -LogName $Log -LogEntryText 'Match found using document ID equals sharing link ID'
                                    }
                                    elseif ($documentId -and $sharingLink.Id -and $sharingLink.Id.ToString().ToLower() -eq $documentId.ToLower()) {
                                        $isMatch = $true
                                        Write-DebugLog -LogName $Log -LogEntryText 'Match found using case-insensitive document ID comparison'
                                    }

                                    if ($isMatch) {
                                        Write-LogEntry -LogName $Log -LogEntryText "Found matching sharing link with ID: $($sharingLink.Id)" -Level 'INFO'

                                        # Store the sharing link URL if we have document details
                                        if ($siteCollectionData[$siteUrl].ContainsKey('DocumentDetails') -and
                                            $siteCollectionData[$siteUrl]['DocumentDetails'].ContainsKey($groupName)) {

                                            # Get the WebUrl property of the sharing link from the link property
                                            $sharingLinkUrl = if ($sharingLink.link -and $sharingLink.link.WebUrl) {
                                                $sharingLink.link.WebUrl
                                            }
                                            else {
                                                'Not found'
                                            }

                                            # Get the expiration date of the sharing link
                                            $expirationDate = 'No expiration'
                                            if ($sharingLink.link -and $sharingLink.link.ExpirationDateTime) {
                                                # Format the expiration date to a readable format
                                                try {
                                                    $expDate = [DateTime]::Parse($sharingLink.link.ExpirationDateTime)
                                                    $expirationDate = $expDate.ToString('yyyy-MM-dd HH:mm:ss')
                                                }
                                                catch {
                                                    $expirationDate = $sharingLink.link.ExpirationDateTime
                                                    Write-DebugLog -LogName $Log -LogEntryText "Could not parse expiration date: $($sharingLink.link.ExpirationDateTime)"
                                                }
                                            }
                                            elseif ($sharingLink.ExpirationDateTime) {
                                                # Alternative location for expiration date
                                                try {
                                                    $expDate = [DateTime]::Parse($sharingLink.ExpirationDateTime)
                                                    $expirationDate = $expDate.ToString('yyyy-MM-dd HH:mm:ss')
                                                }
                                                catch {
                                                    $expirationDate = $sharingLink.ExpirationDateTime
                                                    Write-DebugLog -LogName $Log -LogEntryText "Could not parse expiration date: $($sharingLink.ExpirationDateTime)"
                                                }
                                            }

                                            $siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['SharingLinkUrl'] = $sharingLinkUrl
                                            $siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['ExpirationDate'] = $expirationDate

                                            Write-InfoLog -LogName $Log -LogEntryText "Stored sharing link URL for group $groupName - URL: $sharingLinkUrl, Expiration: $expirationDate"
                                        }

                                        # Use REST API to remove sharing link for list items (based on captured web traffic)
                                        Write-Host '        Attempting to remove sharing link for list item using REST API' -ForegroundColor Cyan
                                        Write-DebugLog -LogName $Log -LogEntryText "Attempting to remove sharing link for list item using REST API: $documentUrl"

                                        try {
                                            # Extract list ID and item ID from the document URL
                                            $listId = ''
                                            $itemId = ''

                                            # Parse item ID from URL
                                            if ($documentUrl -match 'DispForm\.aspx\?ID=(\d+)') {
                                                $itemId = $matches[1]
                                                Write-DebugLog -LogName $Log -LogEntryText "Extracted item ID: $itemId"
                                            }

                                            # Get the list ID by parsing the URL or using PnP to find the list
                                            if ($documentUrl -match '/Lists/([^/]+)/DispForm\.aspx') {
                                                $listName = $matches[1]
                                                Write-DebugLog -LogName $Log -LogEntryText "Extracted list name: $listName"

                                                # Get the list to find its ID
                                                $list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
                                                if ($list) {
                                                    $listId = $list.Id.ToString()
                                                    Write-DebugLog -LogName $Log -LogEntryText "Found list ID: $listId"
                                                }
                                            }

                                            if (-not [string]::IsNullOrWhiteSpace($listId) -and -not [string]::IsNullOrWhiteSpace($itemId) -and $sharingLink.Id) {
                                                # Construct the REST API call similar to the captured web traffic
                                                $restUrl = "$SiteUrl/_api/web/Lists(@a1)/GetItemById(@a2)/UnshareLink?@a1='$($listId.Replace('-','%2D'))'&@a2='$itemId'"

                                                # Determine the link kind. Possible values: https://learn.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.client.sharinglinkkind?view=sharepoint-csom
                                                $linkKind = switch -Regex ($groupName) {
                                                    'AnonymousView' {
                                                        2
                                                    }
                                                    'AnonymousEdit' {
                                                        3
                                                    }
                                                    'OrganizatonView' {
                                                        4
                                                    }
                                                    'OrganizatonEdit' {
                                                        5
                                                    }
                                                }

                                                # Create the request body exactly as captured in web traffic
                                                # The shareId should be the sharing link ID, not the document ID
                                                $requestBody = @{
                                                    linkKind = $linkKind
                                                    shareId  = $sharingLink.Id.ToString()  # Ensure it's a string
                                                }

                                                # Convert to JSON with specific formatting to match captured traffic
                                                $jsonBody = $requestBody | ConvertTo-Json -Compress

                                                Write-DebugLog -LogName $Log -LogEntryText "REST API URL: $restUrl"
                                                Write-DebugLog -LogName $Log -LogEntryText "Request body: $jsonBody"
                                                Write-DebugLog -LogName $Log -LogEntryText "ShareId from sharing link: $($sharingLink.Id)"
                                                Write-DebugLog -LogName $Log -LogEntryText "Sharing link object type: $($sharingLink.GetType().FullName)"

                                                # Try a simpler approach - use the exact same format as captured
                                                $simpleBody = "{`"linkKind`":3,`"shareId`":`"$($sharingLink.Id)`"}"
                                                Write-DebugLog -LogName $Log -LogEntryText "Simple body format: $simpleBody"

                                                # Try using Invoke-RestMethod directly for better control
                                                try {
                                                    Write-DebugLog -LogName $Log -LogEntryText 'Attempting REST call with PnP authentication context'

                                                    # Use PnP's built-in REST method instead of manual token handling
                                                    $response = Invoke-PnPSPRestMethod -Url $restUrl -Method POST -Content $simpleBody

                                                    Write-Host '          Successfully removed sharing link for list item using PnP REST method' -ForegroundColor Green
                                                    Write-InfoLog -LogName $Log -LogEntryText "Successfully removed sharing link for list item using PnP REST method: $documentUrl"
                                                    $linkRemoved = $true
                                                }
                                                catch {
                                                    $pnpRestError = $_.Exception.Message
                                                    Write-DebugLog -LogName $Log -LogEntryText "PnP REST method failed: $pnpRestError"

                                                    # Try manual approach with proper authentication
                                                    try {
                                                        Write-DebugLog -LogName $Log -LogEntryText 'Trying manual REST call with proper authentication'

                                                        # Get the current web context for proper authentication
                                                        $web = Get-PnPWeb
                                                        $context = Get-PnPContext

                                                        # Use CSOM to execute the UnshareLink method directly
                                                        Write-DebugLog -LogName $Log -LogEntryText 'Attempting CSOM UnshareLink method'

                                                        # Get the list by ID
                                                        $list = $context.Web.Lists.GetById($listId)
                                                        $context.Load($list)
                                                        $context.ExecuteQuery()

                                                        # Get the list item
                                                        $listItem = $list.GetItemById($itemId)
                                                        $context.Load($listItem)
                                                        $context.ExecuteQuery()

                                                        # Call UnshareLink directly through CSOM
                                                        $unshareResult = $listItem.UnshareLink(3, $sharingLink.Id)
                                                        $context.ExecuteQuery()

                                                        Write-Host '          Successfully removed sharing link for list item using CSOM UnshareLink' -ForegroundColor Green
                                                        Write-InfoLog -LogName $Log -LogEntryText "Successfully removed sharing link for list item using CSOM UnshareLink: $documentUrl"
                                                        $linkRemoved = $true
                                                    }
                                                    catch {
                                                        $csomError = $_.Exception.Message
                                                        Write-DebugLog -LogName $Log -LogEntryText "CSOM UnshareLink failed: $csomError"

                                                        # Final fallback: try using Remove-PnPFileSharingLink with different parameters
                                                        try {
                                                            Write-DebugLog -LogName $Log -LogEntryText 'Final fallback: trying Remove-PnPFileSharingLink with sharing link ID'

                                                            # Try using the sharing link ID directly
                                                            Remove-PnPFileSharingLink -Identity $sharingLink.Id -Force

                                                            Write-Host '          Successfully removed sharing link using Remove-PnPFileSharingLink fallback' -ForegroundColor Green
                                                            Write-InfoLog -LogName $Log -LogEntryText "Successfully removed sharing link using Remove-PnPFileSharingLink fallback: $documentUrl"
                                                            $linkRemoved = $true
                                                        }
                                                        catch {
                                                            Write-DebugLog -LogName $Log -LogEntryText "All sharing link removal methods failed. Final error: $($_.Exception.Message)"
                                                            Write-Host '          Warning: All sharing link removal methods failed for list item' -ForegroundColor Red
                                                            $linkRemoved = $false
                                                        }
                                                    }
                                                }
                                            }
                                            else {
                                                Write-Host "          Warning: Could not extract required IDs for REST API call (ListID: $listId, ItemID: $itemId, ShareID: $($sharingLink.Id))" -ForegroundColor Yellow
                                                Write-DebugLog -LogName $Log -LogEntryText "Could not extract required IDs for REST API call - ListID: $listId, ItemID: $itemId, ShareID: $($sharingLink.Id)"
                                                $linkRemoved = $false
                                            }
                                        }
                                        catch {
                                            Write-Host "          Warning: REST API call failed for list item: $_" -ForegroundColor Red
                                            Write-ErrorLog -LogName $Log -LogEntryText "REST API call failed for list item: $_"
                                            $linkRemoved = $false
                                        }

                                        break
                                    }
                                }

                                # Since we already handled the link removal above by skipping it,
                                # this fallback won't execute for list items (linkRemoved is already true)
                                if (-not $linkRemoved -and $sharingLinks -and $sharingLinks.Count -gt 0) {
                                    Write-DebugLog -LogName $Log -LogEntryText 'Fallback: This should not execute for list items since linkRemoved is already true'
                                }
                            }
                            else {
                                Write-DebugLog -LogName $Log -LogEntryText "No sharing links found for list item with document ID: $documentId"
                            }
                        }
                        catch {
                            Write-Host "        Error during list item sharing link operations: $_" -ForegroundColor Red
                            Write-ErrorLog -LogName $Log -LogEntryText "List item sharing link operations failed for document ID $documentId : $_"
                        }
                    }
                    else {
                        # This is a regular file - use the original file-based approach
                        Write-DebugLog -LogName $Log -LogEntryText 'Processing as regular file for sharing link removal'

                        try {
                            Write-DebugLog -LogName $Log -LogEntryText "Finding potential sites for document. Available sites: $($allSiteUrls.Count)"
                            $documentSites = $allSiteUrls | Where-Object { $documentUrl.StartsWith($_) } | Sort-Object { $_.Length } -Descending

                            foreach ($documentSite in $documentSites) {
                                Write-DebugLog -LogName $Log -LogEntryText "Checking for document in site: $documentSite"
                                $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
                                if (-not $currentConnection -or $currentConnection.Url -ne $documentSite) {
                                    Connect-PnPOnline -Url $documentSite @connectionParams -ErrorAction Stop
                                }
                                Write-DebugLog -LogName $Log -LogEntryText "Uri: $documentSite"
                                Write-DebugLog -LogName $Log -LogEntryText "Relative Path: $relativePath"
                                $file = Get-PnPFile -Url ([System.Web.HttpUtility]::UrlDecode($relativePath)) -ErrorAction SilentlyContinue
                                if ($file) {
                                    $documentType = 'file'
                                    Write-DebugLog -LogName $Log -LogEntryText 'Document found as a file.'
                                }
                                else {
                                    $file = Get-PnPFolder -Url ([System.Web.HttpUtility]::UrlDecode($relativePath)) -ErrorAction SilentlyContinue
                                    if ($file) {
                                        $documentType = 'folder'
                                        Write-DebugLog -LogName $Log -LogEntryText 'Document found as a folder.'

                                    }
                                    else {
                                        Write-DebugLog -LogName $Log -LogEntryText "Document not found in site: $documentSite"
                                        continue
                                    }
                                }


                                # Get all sharing links for this file using the correct parameter
                                switch ($documentType) {
                                    'file' {
                                        $sharingLinks = Get-PnPFileSharingLink -Identity $relativePath
                                    }
                                    'folder' {
                                        $sharingLinks = Get-PnPFolderSharingLink -Folder $relativePath
                                    }
                                }

                                Write-DebugLog -LogName $Log -LogEntryText "Sharing Links: $($sharingLinks.Count)"
                                # Log the structure of the sharing links for debugging
                                if ($debugLogging -and $sharingLinks -and $sharingLinks.Count -gt 0) {
                                    $firstLink = $sharingLinks[0]
                                    Write-DebugLog -LogName $Log -LogEntryText "Sharing link object properties: $(($firstLink | Get-Member -MemberType Property).Name -join ', ')"

                                    if ($firstLink.link) {
                                        Write-DebugLog -LogName $Log -LogEntryText "Link property exists. Link properties: $(($firstLink.link | Get-Member -MemberType Property).Name -join ', ')"
                                        if ($firstLink.link.WebUrl) {
                                            Write-DebugLog -LogName $Log -LogEntryText "WebUrl found: $($firstLink.link.WebUrl)"
                                        }
                                    }
                                    else {
                                        Write-DebugLog -LogName $Log -LogEntryText "Link property doesn't exist or is null"
                                    }
                                }
                                else {
                                    Write-DebugLog -LogName $Log -LogEntryText "No sharing links found for $relativePath"
                                }

                                foreach ($sharingLink in $sharingLinks) {
                                    # Try to match the sharing link with our group
                                    if ($sharingLink.Id -and $groupName -like "*$($sharingLink.Id)*") {
                                        Write-LogEntry -LogName $Log -LogEntryText "Found matching sharing link with ID: $($sharingLink.Id)" -Level 'INFO'

                                        # Store the sharing link URL if we have document details
                                        if ($siteCollectionData[$SiteUrl].ContainsKey('DocumentDetails') -and
                                            $siteCollectionData[$SiteUrl]['DocumentDetails'].ContainsKey($groupName)) {

                                            # Get the WebUrl property of the sharing link from the link property
                                            $sharingLinkUrl = if ($sharingLink.link -and $sharingLink.link.WebUrl) {
                                                $sharingLink.link.WebUrl
                                            }
                                            else {
                                                'Not found'
                                            }

                                            # Get the expiration date of the sharing link
                                            $expirationDate = 'No expiration'
                                            if ($sharingLink.link -and $sharingLink.link.ExpirationDateTime) {
                                                # Format the expiration date to a readable format
                                                try {
                                                    $expDate = [DateTime]::Parse($sharingLink.link.ExpirationDateTime)
                                                    $expirationDate = $expDate.ToString('yyyy-MM-dd HH:mm:ss')
                                                }
                                                catch {
                                                    $expirationDate = $sharingLink.link.ExpirationDateTime
                                                    Write-DebugLog -LogName $Log -LogEntryText "Could not parse expiration date: $($sharingLink.link.ExpirationDateTime)"
                                                }
                                            }
                                            elseif ($sharingLink.ExpirationDateTime) {
                                                # Alternative location for expiration date
                                                try {
                                                    $expDate = [DateTime]::Parse($sharingLink.ExpirationDateTime)
                                                    $expirationDate = $expDate.ToString('yyyy-MM-dd HH:mm:ss')
                                                }
                                                catch {
                                                    $expirationDate = $sharingLink.ExpirationDateTime
                                                    Write-DebugLog -LogName $Log -LogEntryText "Could not parse expiration date: $($sharingLink.ExpirationDateTime)"
                                                }
                                            }

                                            $siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['SharingLinkUrl'] = $sharingLinkUrl
                                            $siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['ExpirationDate'] = $expirationDate

                                            Write-InfoLog -LogName $Log -LogEntryText "Stored sharing link URL for group $groupName - URL: $sharingLinkUrl, Expiration: $expirationDate"
                                        }

                                        # Remove the sharing link using the file URL and sharing link ID
                                        switch ($documentType) {
                                            'file' {
                                                Remove-PnPFileSharingLink -FileUrl $relativePath -Id $sharingLink.Id -Force
                                            }
                                            'folder' {
                                                Remove-PnPFolderSharingLink -Folder $relativePath -Identity $sharingLink.Id -Force
                                            }
                                        }

                                        Write-Host '        Successfully removed sharing link using PnP methods' -ForegroundColor Green
                                        Write-InfoLog -LogName $Log -LogEntryText "Successfully removed sharing link with ID: $($sharingLink.Id)"
                                        $linkRemoved = $true
                                        break
                                    }
                                }
                            }
                        }
                        catch {
                            Write-LogEntry -LogName $Log -LogEntryText "PnP sharing link methods failed: $_" -Level 'DEBUG'
                            # Fall through to alternative methods
                        }
                    }

                    # Return the result
                    return $linkRemoved
                } -Operation 'Unshare link using PnP and CSOM methods'

                # Set the result from the script block
                $sharingLinkRemoved = $result

                if (-not $sharingLinkRemoved) {
                    Write-ErrorLog -LogName $Log -LogEntryText "All sharing link removal methods failed for group: $groupName"
                }
            }
            catch {
                Write-ErrorLog -LogName $Log -LogEntryText "All sharing link removal methods failed: $_"
            }
        }
        else {
            Write-DebugLog -LogName $Log -LogEntryText 'Insufficient document details to attempt to unshare using PnP'
        }

        # Always try to remove the empty group after removing members and sharing links
        Write-LogEntry -LogName $Log -LogEntryText "Attempting to remove sharing group: $groupName" -Level 'INFO'
        $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if (-not $currentConnection -or $currentConnection.Url -ne $siteUrl) {
            Connect-PnPOnline -Url $siteUrl @connectionParams -ErrorAction Stop
        }

        try {
            Invoke-WithThrottleHandling -ScriptBlock {
                Write-DebugLog -LogName $Log -LogEntryText "Looking up group with name $groupName from $siteUrl"
                # Try Force parameter first, fallback to no confirmation parameter
                try {
                    # First check if group still exists
                    $groupCheck = Get-PnPGroup $groupName -ErrorAction SilentlyContinue
                    if ($groupCheck) {
                        Remove-PnPGroup -Identity $groupCheck.Id -Force
                        Write-Host "      Successfully removed sharing group: $groupName" -ForegroundColor Green
                        Write-InfoLog -LogName $Log -LogEntryText "Successfully removed sharing group: $groupName"
                    }
                    else {
                        Write-LogEntry -LogName $Log -LogEntryText "Group $groupName no longer exists, may have already been removed" -Level 'INFO'
                        Write-Host '      Group no longer exists (may have already been removed)' -ForegroundColor Yellow
                    }
                }
                catch {
                    # Fallback if Force parameter is not supported
                    $groupCheck = Get-PnPGroup $groupName -ErrorAction SilentlyContinue
                    if ($groupCheck) {
                        Remove-PnPGroup -Identity $groupCheck.Id
                        Write-Host "      Successfully removed empty sharing group: $groupName" -ForegroundColor Green
                        Write-InfoLog -LogName $Log -LogEntryText "Successfully removed sharing group: $groupName"
                    }
                    else {
                        Write-LogEntry -LogName $Log -LogEntryText "Group $groupName no longer exists during fallback removal" -Level 'INFO'
                        Write-Host '      Group no longer exists (fallback check)' -ForegroundColor Yellow
                    }
                }
            } -Operation "Remove sharing group $groupName"
        }
        catch {
            Write-Host "      Warning: Could not remove sharing group $groupName : $_" -ForegroundColor Red
            Write-ErrorLog -LogName $Log -LogEntryText "Final attempt failed to remove sharing group $groupName : $_"
        }

        # Update the link removal status in site collection data
        Update-LinkRemovalStatus -SiteUrl $siteUrl -SharingGroupName $groupName -WasRemoved $true

    }
    catch {
        Write-Host "      Warning: Error during sharing link removal for $groupName : $_" -ForegroundColor Red
        Write-ErrorLog -LogName $Log -LogEntryText "Error during sharing link removal for $groupName : $_"

        # Update status as failed for this group
        Update-LinkRemovalStatus -SiteUrl $siteUrl -SharingGroupName $groupName -WasRemoved $false
    }
}

# ----------------------------------------------
# Function to detect and parse script's CSV output for Organization links
# ----------------------------------------------
function Test-AndParseScriptCsvOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string] $FilePath
    )

    try {
        # Read the first few lines to check the header format
        $firstLine = Get-Content -Path $FilePath -TotalCount 1

        # Check if this looks like our script's CSV output format
        $expectedHeaders = @('Site URL', 'Site Owner', 'IB Mode', 'IB Segment', 'Site Template', 'Sharing Group Name', 'Sharing Link Members', 'File URL', 'File Owner', 'IsTeamsConnected', 'SharingCapability', 'Last Content Modified', 'Search Status', 'Link Removed')

        if ($firstLine -and $firstLine.Contains('Sharing Group Name')) {
            Write-Host "Detected script's CSV output format - will process Organization sharing links only" -ForegroundColor Cyan
            Write-InfoLog -LogName $Log -LogEntryText "Input file detected as script's CSV output format"

            # Import the full CSV
            $csvData = Import-Csv -Path $FilePath

            # Filter for Organization sharing links only
            $organizationEntries = $csvData | Where-Object {
                $_.'Sharing Group Name' -like '*Organization*' -and
                -not [string]::IsNullOrWhiteSpace($_.'Site URL')
            }

            if ($organizationEntries.Count -eq 0) {
                Write-Host 'No Organization sharing links found in the input CSV file' -ForegroundColor Yellow
                Write-InfoLog -LogName $Log -LogEntryText 'No Organization sharing links found in input CSV'
                return @{
                    IsScriptOutput    = $true
                    Sites             = @()
                    OrganizationLinks = @{
                    }
                }
            }

            # Group by Site URL to get unique sites
            $siteGroups = $organizationEntries | Group-Object 'Site URL'

            # Create sites collection for processing
            $sitesToProcess = [System.Collections.ArrayList]::new()
            $orgLinksData = @{
            }

            foreach ($siteGroup in $siteGroups) {
                $siteUrl = $siteGroup.Name
                $sitesToProcess.Add([PSCustomObject]@{ URL = $siteUrl }) | Out-Null

                # Store Organization sharing group details for this site
                $orgLinksData[$siteUrl] = @{
                    Groups               = [System.Collections.ArrayList]::new()
                    HasOrganizationLinks = $true
                }

                foreach ($entry in $siteGroup.Group) {
                    $orgLinksData[$siteUrl].Groups.Add(@{
                            GroupName = $entry.'Sharing Group Name'
                            Members   = $entry.'Sharing Link Members'
                            FileUrl   = $entry.'File URL'
                            FileOwner = $entry.'File Owner'
                        }) | Out-Null
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
            Write-Host 'Input file appears to be a simple site URL list' -ForegroundColor Yellow
            Write-InfoLog -LogName $Log -LogEntryText 'Input file detected as simple site URL list'

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

# ----------------------------------------------
# Function to search for documents using Microsoft Graph API
# ----------------------------------------------
function Search-DocumentViaGraphAPI {
    param(
        [Parameter(Mandatory = $true)]
        [string] $DocumentId,
        [Parameter(Mandatory = $true)]
        [string] $SearchRegion,
        [Parameter(Mandatory = $false)]
        [string] $LogContext = 'Document search'
    )

    $result = @{
        Found                = $false
        DocumentUrl          = ''
        DocumentOwner        = ''
        ItemType             = ''
        DocumentLastModified = ''
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
            'Authorization' = "Bearer $graphToken"
            'Content-Type'  = 'application/json'
        }

        $searchUrl = 'https://graph.microsoft.com/v1.0/search/query'
        $itemFound = $false

        # First, try searching as driveItem (files in document libraries)
        $driveItemSearchQuery = @{
            requests = @(
                @{
                    entityTypes               = @('driveItem')
                    query                     = @{
                        queryString = "UniqueID:$DocumentId"
                    }
                    from                      = 0
                    size                      = 25
                    sharePointOneDriveOptions = @{
                        includeContent = 'sharedContent,privateContent'
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
                $result.ItemType = 'driveItem'

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

                if ($resource.fileSystemInfo.lastModifiedDateTime) {
                    $result.DocumentLastModified = $resource.fileSystemInfo.lastModifiedDateTime
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
                        entityTypes               = @('listItem')
                        query                     = @{
                            queryString = "UniqueID:$DocumentId"
                        }
                        from                      = 0
                        size                      = 25
                        sharePointOneDriveOptions = @{
                            includeContent = 'sharedContent,privateContent'
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
                    $result.ItemType = 'listItem'

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
function Get-SharingLinkUrls {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,

        [bool]$IgnoreFlexibleLinkGroups = $false
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
            Get-PnPGroup | Where-Object { $_.Title -like 'SharingLinks*' }
        } -Operation "Get sharing groups for $SiteUrl"

        if ($sharingGroups.Count -eq 0) {
            Write-DebugLog -LogName $Log -LogEntryText "No sharing groups found on site: $SiteUrl"
            return
        }

        Write-Host "    Found $($sharingGroups.Count) sharing groups" -ForegroundColor Green
        Write-InfoLog -LogName $Log -LogEntryText "Found $($sharingGroups.Count) sharing groups on site: $SiteUrl"

        if ($ignoreFlexibleLinkGroups) {
            $sharingGroups = $sharingGroups | Where-Object { $_.Title -notlike '*Flexible*' }
            Write-Host "    Found $($sharingGroups.Count) sharing groups which are not 'Flexible'" -ForegroundColor Green
            Write-InfoLog -LogName $Log -LogEntryText "Found $($sharingGroups.Count) sharing groups on site which are not 'Flexible': $SiteUrl"
        }

        foreach ($group in $sharingGroups) {
            $groupName = $group.Title

            # Extract document ID from group name
            $documentId = ''
            if ($groupName -match 'SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.') {
                $documentId = $matches[1]
                Write-DebugLog -LogName $Log -LogEntryText "Processing sharing group: $groupName with document ID: $documentId"

                # Check if we have document details for this group
                if ($siteCollectionData[$SiteUrl].ContainsKey('DocumentDetails') -and
                    $siteCollectionData[$SiteUrl]['DocumentDetails'].ContainsKey($groupName) -and
                    -not [string]::IsNullOrWhiteSpace($siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['DocumentUrl'])) {

                    $docUrl = $siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['DocumentUrl']

                    # Try to get sharing links using the document ID directly (works for both files and list items)
                    try {
                        $sharingLinks = $null
                        $sharingLinkUrl = 'Not found'
                        $expirationDate = 'No expiration'

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
                                    'Not found'
                                }

                                # Check for members in the sharing link itself (GrantedToIdentitiesV2, GrantedToV2)
                                # This is where SharePoint sometimes stores the actual users with access
                                $sharingLinkMembers = @()

                                # Try GrantedToIdentitiesV2 first (newer format)
                                if ($matchingLink.GrantedToIdentitiesV2 -and $matchingLink.GrantedToIdentitiesV2.Count -gt 0) {
                                    Write-DebugLog -LogName $Log -LogEntryText "Found $($matchingLink.GrantedToIdentitiesV2.Count) members in GrantedToIdentitiesV2 for sharing link"
                                    foreach ($identity in $matchingLink.GrantedToIdentitiesV2) {
                                        if ($identity.User) {
                                            $memberEmail = if ($identity.User.Email) {
                                                $identity.User.Email
                                            }
                                            else {
                                                ''
                                            }
                                            $memberDisplayName = if ($identity.User.DisplayName) {
                                                $identity.User.DisplayName
                                            }
                                            else {
                                                $memberEmail
                                            }
                                            $memberLoginName = if ($identity.User.Id) {
                                                $identity.User.Id
                                            }
                                            else {
                                                $memberEmail
                                            }

                                            Write-DebugLog -LogName $Log -LogEntryText "  GrantedToIdentitiesV2 member: DisplayName='$memberDisplayName', Email='$memberEmail', Id='$memberLoginName'"

                                            # Add this member to the site collection data if not already present
                                            $existingMember = $siteCollectionData[$SiteUrl]['SP Users'] | Where-Object {
                                                $_.AssociatedSPGroup -eq $groupName -and
                                                ($_.Name -eq $memberLoginName -or $_.Email -eq $memberEmail)
                                            }

                                            if (-not $existingMember) {
                                                Write-DebugLog -LogName $Log -LogEntryText "  Adding sharing link member to site data: Group='$groupName', LoginName='$memberLoginName', DisplayName='$memberDisplayName', Email='$memberEmail'"

                                                $userObject = [PSCustomObject]@{
                                                    AssociatedSPGroup = $groupName
                                                    Name              = $memberLoginName
                                                    Title             = $memberDisplayName
                                                    Email             = $memberEmail
                                                }
                                                $siteCollectionData[$SiteUrl]['SP Users'].Add($userObject)
                                            }
                                        }
                                    }
                                }
                                # Fallback to GrantedToV2 (older format)
                                elseif ($matchingLink.GrantedToV2 -and $matchingLink.GrantedToV2.Count -gt 0) {
                                    Write-DebugLog -LogName $Log -LogEntryText "Found $($matchingLink.GrantedToV2.Count) members in GrantedToV2 for sharing link"
                                    foreach ($grantee in $matchingLink.GrantedToV2) {
                                        if ($grantee.User) {
                                            $memberEmail = if ($grantee.User.Email) {
                                                $grantee.User.Email
                                            }
                                            else {
                                                ''
                                            }
                                            $memberDisplayName = if ($grantee.User.DisplayName) {
                                                $grantee.User.DisplayName
                                            }
                                            else {
                                                $memberEmail
                                            }
                                            $memberLoginName = if ($grantee.User.Id) {
                                                $grantee.User.Id
                                            }
                                            else {
                                                $memberEmail
                                            }

                                            Write-DebugLog -LogName $Log -LogEntryText "  GrantedToV2 member: DisplayName='$memberDisplayName', Email='$memberEmail', Id='$memberLoginName'"

                                            # Add this member to the site collection data if not already present
                                            $existingMember = $siteCollectionData[$SiteUrl]['SP Users'] | Where-Object {
                                                $_.AssociatedSPGroup -eq $groupName -and
                                                ($_.Name -eq $memberLoginName -or $_.Email -eq $memberEmail)
                                            }

                                            if (-not $existingMember) {
                                                Write-DebugLog -LogName $Log -LogEntryText "  Adding sharing link member to site data: Group='$groupName', LoginName='$memberLoginName', DisplayName='$memberDisplayName', Email='$memberEmail'"

                                                $userObject = [PSCustomObject]@{
                                                    AssociatedSPGroup = $groupName
                                                    Name              = $memberLoginName
                                                    Title             = $memberDisplayName
                                                    Email             = $memberEmail
                                                }
                                                $siteCollectionData[$SiteUrl]['SP Users'].Add($userObject)
                                            }
                                        }
                                    }
                                }
                                else {
                                    Write-DebugLog -LogName $Log -LogEntryText 'No members found in GrantedToIdentitiesV2 or GrantedToV2 for sharing link - this may be an anonymous/anyone link'
                                }

                                # Get the expiration date of the sharing link
                                if ($matchingLink.link -and $matchingLink.link.ExpirationDateTime) {
                                    # Format the expiration date to a readable format
                                    try {
                                        $expDate = [DateTime]::Parse($matchingLink.link.ExpirationDateTime)
                                        $expirationDate = $expDate.ToString('yyyy-MM-dd HH:mm:ss')
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
                                        $expirationDate = $expDate.ToString('yyyy-MM-dd HH:mm:ss')
                                    }
                                    catch {
                                        $expirationDate = $matchingLink.ExpirationDateTime
                                        Write-DebugLog -LogName $Log -LogEntryText "Could not parse expiration date: $($matchingLink.ExpirationDateTime)"
                                    }
                                }

                                # Store the results
                                $siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['SharingLinkUrl'] = $sharingLinkUrl
                                $siteCollectionData[$SiteUrl]['DocumentDetails'][$groupName]['ExpirationDate'] = $expirationDate

                                #Write-Host "      Found sharing link URL for group: $groupName" -ForegroundColor Green
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
$organizationLinksProcessedCount = 0
$anyoneLinksProcessedCount = 0

# Display script mode information again before starting site processing
Write-Host ''
Write-Host '======================================================' -ForegroundColor $(if ($convertOrganizationLinks) {
        'Yellow'
    }
    else {
        'Cyan'
    })
Write-Host "SCRIPT MODE: $scriptMode" -ForegroundColor $(if ($convertOrganizationLinks) {
        'Yellow'
    }
    else {
        'Cyan'
    })

if ($convertOrganizationLinks) {
    Write-Host '  - Organization sharing links will be CONVERTED to direct permissions' -ForegroundColor Yellow
    Write-Host '  - Organization sharing links will be REMOVED after converting users' -ForegroundColor Yellow
    Write-Host '  - Empty Organization sharing groups will be cleaned up automatically' -ForegroundColor Yellow
    Write-Host '  - Flexible sharing links will be PRESERVED (not modified)' -ForegroundColor Green
    Write-Host "  - Results will be saved to: $sharingLinksOutputFile" -ForegroundColor Yellow
    if ($removeAnyoneLinks) {
        Write-Host '  - Anonymous sharing links will be REMOVED' -ForegroundColor Yellow
    }
}
else {
    Write-Host '  - Only DETECTING and INVENTORYING sharing links' -ForegroundColor Cyan
    Write-Host '  - NO modifications will be made to permissions or sharing links' -ForegroundColor Cyan
    Write-Host "  - Results will be saved to: $sharingLinksOutputFile" -ForegroundColor Cyan
    if ($ignoreFlexibleLinkGroups) {
        Write-Host '  - Flexible link groups and links will NOT be included in the output' -ForegroundColor Cyan
    }
}
Write-Host '======================================================' -ForegroundColor $(if ($convertOrganizationLinks) {
        'Yellow'
    }
    else {
        'Cyan'
    })
Write-Host ''
Write-InfoLog -LogName $Log -LogEntryText "Starting to process $totalSites sites in $scriptMode mode"

foreach ($site in $sites) {
    $processedCount++
    $siteUrl = ''

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
            # Get Site Properties using SharePoint Admin connection
            $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
            if (-not $currentConnection -or $currentConnection.Url -ne $adminUrl) {
                Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop
            }
            $siteProperties = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPTenantSite -Identity $siteUrl
            } -Operation "Get site properties for $siteUrl"

            # Connect back to the site for group processing
            $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
            if (-not $currentConnection -or $currentConnection.Url -ne $SiteUrl) {
                Connect-PnPOnline -Url $SiteUrl @connectionParams -ErrorAction Stop
            }

            # Initialize site data
            Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteProperties

            $token = Get-PnPGraphTokenCompatible
            $graphBeta = 'https://graph.microsoft.com/beta'
            $rootsEndpoint = "$GraphBeta/sites?filter=siteCollection/root ne null&select=webUrl, siteCollection"
            $SiteLocationData = Invoke-RestMethod -Uri $rootsEndpoint -Headers @{Authorization = "Bearer $token" } -Method GET
            $SearchRegion = $SiteLocationData.value.siteCollection.dataLocationCode
            Write-DebugLog -LogName $Log -LogEntryText "Determined site region: $SearchRegion"

            # Get all groups for this site
            $spGroups = Invoke-WithThrottleHandling -ScriptBlock {
                Get-PnPGroup
            } -Operation "Get groups for site $siteUrl"

            if ($ignoreFlexibleLinkGroups) {
                $spGroups = $spGroups | Where-Object { $_.Title -notlike '*Flexible*' }
                Write-Host " Found $($spGroups.Count) groups on the site which are not 'Flexible'" -ForegroundColor Green
                Write-InfoLog -LogName $Log -LogEntryText "Found $($spGroups.Count) groups on site which are not 'Flexible': $SiteUrl"
            }

            foreach ($spGroup in $spGroups) {
                $spGroupName = $spGroup.Title

                # Update site data with group information
                Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteProperties -SPGroupName $spGroupName

                # Get users in each group with enhanced external user support
                $spUsers = Invoke-WithThrottleHandling -ScriptBlock {
                    # Try the standard approach first
                    $standardUsers = Get-PnPGroupMember -Identity $spGroup.Id -ErrorAction SilentlyContinue

                    # For sharing groups, also try alternative approaches to catch external users
                    if ($spGroupName -like 'SharingLinks*') {
                        try {
                            # Try using CSOM to get all users including external ones
                            $ctx = Get-PnPContext
                            $group = $ctx.Web.SiteGroups.GetById($spGroup.Id)
                            $users = $group.Users
                            $ctx.Load($users)
                            $ctx.ExecuteQuery()

                            # Convert CSOM users to PnP format for consistency
                            $csomUsers = [System.Collections.ArrayList]::new()
                            foreach ($user in $users) {
                                $csomUsers.Add([PSCustomObject]@{
                                        Id            = $user.Id
                                        LoginName     = $user.LoginName
                                        Title         = $user.Title
                                        Email         = $user.Email
                                        PrincipalType = $user.PrincipalType
                                    })
                            }

                            # Combine standard and CSOM results, removing duplicates by LoginName
                            $allUsers = @($standardUsers) + @($csomUsers) | Group-Object LoginName | ForEach-Object { $_.Group[0] }

                            Write-DebugLog -LogName $Log -LogEntryText "Group '$spGroupName': Standard method found $($standardUsers.Count) users, CSOM found $($csomUsers.Count) users, combined unique: $($allUsers.Count) users"

                            return $allUsers
                        }
                        catch {
                            Write-DebugLog -LogName $Log -LogEntryText "CSOM fallback failed for group '$spGroupName': $_. Using standard results only."
                            return $standardUsers
                        }
                    }
                    else {
                        return $standardUsers
                    }
                } -Operation "Get members for group $spGroupName"

                # Debug: Log the number of users found and their basic info
                if ($spGroupName -like 'SharingLinks*') {
                    Write-DebugLog -LogName $Log -LogEntryText "Sharing group '$spGroupName' has $($spUsers.Count) members"

                    if ($spUsers.Count -gt 0) {
                        foreach ($debugUser in $spUsers) {
                            Write-DebugLog -LogName $Log -LogEntryText "  Member found - LoginName: '$($debugUser.LoginName)', Title: '$($debugUser.Title)', Email: '$($debugUser.Email)', PrincipalType: '$($debugUser.PrincipalType)'"
                        }
                    }
                    else {
                        Write-DebugLog -LogName $Log -LogEntryText "  No members found in sharing group '$spGroupName'"
                    }
                }

                foreach ($spUser in $spUsers) {
                    # Enhanced null checking - external users might have different property patterns
                    $hasValidLoginName = -not [string]::IsNullOrWhiteSpace($spUser.LoginName)
                    $hasValidId = $spUser.Id -ne $null -and $spUser.Id -gt 0

                    if ($spUser -and ($hasValidLoginName -or $hasValidId)) {
                        # Debug: Show what we're storing for sharing links groups
                        if ($spGroupName -like 'SharingLinks*') {
                            Write-DebugLog -LogName $Log -LogEntryText "  Storing member for '$spGroupName': LoginName='$($spUser.LoginName)', Title='$($spUser.Title)', Email='$($spUser.Email)', Id='$($spUser.Id)', PrincipalType='$($spUser.PrincipalType)'"
                        }

                        # Use LoginName as primary identifier, fallback to Title if LoginName is empty (for some external users)
                        $userIdentifier = if (-not [string]::IsNullOrWhiteSpace($spUser.LoginName)) {
                            $spUser.LoginName
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace($spUser.Title)) {
                            $spUser.Title  # Fallback for edge cases
                        }
                        else {
                            "User_$($spUser.Id)"  # Last resort fallback
                        }

                        Update-SiteCollectionData -SiteUrl $siteUrl -SiteProperties $siteProperties -AssociatedSPGroup $spGroupName -SPUserName $userIdentifier -SPUserTitle $spUser.Title -SPUserEmail $spUser.Email
                    }
                    else {
                        # Debug: Log why we're skipping this user with more detail
                        if ($spGroupName -like 'SharingLinks*') {
                            Write-DebugLog -LogName $Log -LogEntryText "  Skipping member in '$spGroupName' - spUser is null: $($spUser -eq $null), LoginName: '$($spUser.LoginName)', Id: '$($spUser.Id)', Title: '$($spUser.Title)'"
                        }
                    }
                }

                # Extract document information from sharing groups
                if ($spGroupName -like 'SharingLinks*') {
                    try {
                        # Extract document ID from sharing group name
                        if ($spGroupName -match 'SharingLinks\.([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.') {
                            $documentId = $matches[1]
                            Write-DebugLog -LogName $Log -LogEntryText "Extracted document ID: $documentId from sharing group: $spGroupName"
                            $sharingType = 'Unknown'
                            $documentUrl = ''
                            $documentOwner = ''
                            $documentItemType = ''
                            $searchStatus = 'Not Searched'  # Track if document was found in search results

                            # Determine sharing type from group name
                            if ($spGroupName -like '*OrganizationView*') {
                                $sharingType = 'OrganizationView'
                            }
                            elseif ($spGroupName -like '*OrganizationEdit*') {
                                $sharingType = 'OrganizationEdit'
                            }
                            elseif ($spGroupName -like '*AnonymousEdit*') {
                                $sharingType = 'AnonymousEdit'
                            }
                            elseif ($spGroupName -like '*AnonymousView*') {
                                $sharingType = 'AnonymousView'
                            }

                            # Try to find the document using Microsoft Graph
                            try {
                                $graphToken = Invoke-WithThrottleHandling -ScriptBlock {
                                    Get-PnPGraphTokenCompatible
                                } -Operation 'Get Graph access token (version-compatible) for document search'

                                if ($graphToken) {
                                    $headers = @{
                                        'Authorization' = "Bearer $graphToken"
                                        'Content-Type'  = 'application/json'
                                    }

                                    # Try to find the document via Microsoft Graph search using the document ID
                                    $searchResult = Search-DocumentViaGraphAPI -DocumentId $documentId -SearchRegion $searchRegion -LogContext 'Main processing loop - document search'

                                    Write-DebugLog -LogName $Log -LogEntryText "Search result for document ID $documentId - Found: $($searchResult.Found), URL: '$($searchResult.DocumentUrl)', Owner: '$($searchResult.DocumentOwner)', Type: '$($searchResult.ItemType)'"

                                    if ($searchResult.Found) {
                                        $searchStatus = 'Found'
                                        if ($searchResult.DocumentUrl) {
                                            $documentUrl = $searchResult.DocumentUrl
                                        }

                                        if ($searchResult.DocumentOwner) {
                                            $documentOwner = $searchResult.DocumentOwner
                                        }

                                        if ($searchResult.ItemType) {
                                            $documentItemType = $searchResult.ItemType
                                        }

                                        if ($searchResult.DocumentLastModified) {
                                            $documentLastModified = $searchResult.DocumentLastModified
                                        }
                                    }
                                    else {
                                        $searchStatus = 'Not Found in Search'
                                        # Set default values to indicate the file was not searchable
                                        $documentUrl = 'Not Searchable'
                                        $documentOwner = 'Not Searchable'
                                        $documentLastModified = 'Not Searchable'
                                        $documentItemType = 'Not Searchable'
                                    }
                                }
                                else {
                                    Write-LogEntry -LogName $Log -LogEntryText 'Unable to get Graph access token for document search.' -Level 'ERROR'
                                    $searchStatus = 'Search Error'
                                    $documentUrl = 'Search Error'
                                    $documentOwner = 'Search Error'
                                    $documentLastModified = 'Search Error'
                                    $documentItemType = 'Search Error'
                                }
                            }
                            catch {
                                Write-ErrorLog -LogName $Log -LogEntryText "Error searching for document via Graph API: ${_}"
                                $searchStatus = 'Search Error'
                                $documentUrl = 'Search Error'
                                $documentOwner = 'Search Error'
                                $documentLastModified = 'Search Error'
                                $documentItemType = 'Search Error'
                            }

                            # Store the sharing link information
                            if (-not $siteCollectionData[$siteUrl].ContainsKey('DocumentDetails')) {
                                $siteCollectionData[$siteUrl]['DocumentDetails'] = @{
                                }
                            }

                            $siteCollectionData[$siteUrl]['DocumentDetails'][$spGroupName] = @{
                                'DocumentId'           = $documentId
                                'SharingType'          = $sharingType
                                'DocumentUrl'          = $documentUrl
                                'DocumentOwner'        = $documentOwner
                                'DocumentLastModified' = $documentLastModified
                                'DocumentItemType'     = $documentItemType
                                'SearchStatus'         = $searchStatus
                                'SharedOn'             = $siteUrl
                                'SharingLinkUrl'       = '' # Will be populated when processing sharing links
                                'ExpirationDate'       = '' # Will be populated when processing sharing links
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
            if ($siteCollectionData[$siteUrl]['Has Sharing Links']) {
                $sitesWithSharingLinksCount++

                # Collect sharing link URLs for all sites, whether in detection or remediation mode
                Get-SharingLinkUrls -SiteUrl $siteUrl -IgnoreFlexibleLinkGroups $ignoreFlexibleLinkGroups

                # Convert Organization sharing links to direct permissions if enabled
                if ($convertOrganizationLinks) {
                    Convert-OrganizationSharingLinks -SiteUrl $siteUrl
                    $organizationLinksProcessedCount++
                }

                if ($removeAnyoneLinks) {
                    Remove-AnyoneSharingLinks -SiteUrl $siteUrl
                    $anyoneLinksProcessedCount++
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
Write-Host 'Consolidating results...' -ForegroundColor Green

# No incremental file generation - only focus on sharing links output
if ($sitesWithSharingLinksCount -gt 0) {
    Write-Host "Found $sitesWithSharingLinksCount site collections with sharing links" -ForegroundColor Green
    Write-Host "Sharing links data written to: $sharingLinksOutputFile" -ForegroundColor Green
    Write-InfoLog -LogName $Log -LogEntryText "Total sites with sharing links: $sitesWithSharingLinksCount"

    if ($convertOrganizationLinks) {
        Write-Host "Processed Organization sharing links on $organizationLinksProcessedCount sites" -ForegroundColor Green
        Write-Host '  Mode: REMEDIATION - Organization links were converted to direct permissions and removed (flexible links preserved)' -ForegroundColor Yellow
        Write-Host '  Group cleanup: ENABLED - Empty sharing groups were cleaned up' -ForegroundColor Yellow
        Write-InfoLog -LogName $Log -LogEntryText "Processed Organization sharing links on $organizationLinksProcessedCount sites in REMEDIATION mode (convertOrganizationLinks=$convertOrganizationLinks, cleanupCorruptedSharingGroups=$cleanupCorruptedSharingGroups)"
    }

    if ($removeAnyoneLinks) {
        Write-Host "Processed Anyone sharing links on $anyoneLinksProcessedCount sites" -ForegroundColor Green
        Write-Host '  Mode: REMEDIATION - Anyone links were removed' -ForegroundColor Yellow
        Write-InfoLog -LogName $Log -LogEntryText "Processed Anyone sharing links on $anyoneLinksProcessedCount sites in REMEDIATION mode (removeAnyoneLinks=$removeAnyoneLinks)"
    }
}
else {
    Write-Host 'No site collections with sharing links found.' -ForegroundColor Yellow
    Write-InfoLog -LogName $Log -LogEntryText 'No site collections with sharing links found.'
}

# ----------------------------------------------
# Disconnect and finish
# ----------------------------------------------
Disconnect-PnPOnline
Write-InfoLog -LogName $Log -LogEntryText 'Script finished.'
Write-Host "Script finished. Log file located at: $log" -ForegroundColor Green
