<#
.SYNOPSIS
    SharePoint Document Permission Checker - Validates user access to specific documents using Microsoft Graph and SharePoint REST APIs.

.DESCRIPTION
    This script checks whether a specified user has access to a particular SharePoint document by examining:
    - Direct user permissions on the document
    - SharePoint group memberships
    - Microsoft 365/Entra ID group memberships  
    - "Everyone except external users" permissions
    - Inherited permissions from parent libraries
    
    The script uses certificate-based authentication to connect to both Microsoft Graph API and SharePoint REST API
    to provide comprehensive permission analysis with validation against SharePoint roleassignments for accuracy.

.PARAMETER appID
    The Entra ID Application ID for certificate-based authentication

.PARAMETER thumbprint  
    The certificate thumbprint for authentication

.PARAMETER tenant
    The Tenant ID for authentication

.PARAMETER t
    The tenant name (e.g., 'contoso' for contoso.sharepoint.com)

.PARAMETER siteUrl
    The full SharePoint site URL to check (e.g., 'https://tenant.sharepoint.com/sites/sitename')

.PARAMETER userToCheck
    The user principal name (email) of the user to check permissions for

.PARAMETER documentUrl
    The full SharePoint document URL including path and filename

.PARAMETER debug
    Enable detailed debug output for troubleshooting ($true/$false)

.EXAMPLE
    # Configure the script variables at the top and run
    $siteUrl = 'https://contoso.sharepoint.com/sites/team'
    $userToCheck = 'user@contoso.com'  
    $documentUrl = 'https://contoso.sharepoint.com/sites/team/Shared%20Documents/file.pdf'
    
.OUTPUTS
    CSV file: Contains permission analysis results with columns for SiteName, URL, DocumentName, DocumentURL, User, Owner, AccessType
    Log file: Detailed logging of all operations and API calls

.NOTES
    File Name      : SPOPerms-Checker.ps1
    Author         : Mike Lee
    Date           : 12/3/25
    Prerequisite   : 
    - PowerShell 5.1 or later (7.5 Preferred)
    - Certificate installed in certificate store for app authentication
    - Entra ID app registration with appropriate permissions:
      - Microsoft Graph: Sites.Read.All, Group.Read.All, User.Read.All
      - SharePoint: Sites.FullControl.All (for REST API access)
    
    Version History:
    - Supports both unique document permissions and inherited permissions
    - Validates Graph API results against SharePoint REST API roleassignments
    - Handles "Everyone except external users" permission detection and validation
    - Provides comprehensive group membership checking (SharePoint and Entra ID groups)

.FUNCTIONALITY
    Authentication:
    - Uses certificate-based authentication for secure API access
    - Acquires separate tokens for Graph API and SharePoint REST API
    
    Permission Analysis:
    - Retrieves document by direct URL for accuracy
    - Checks document-level permissions via Graph API
    - Validates permissions via SharePoint REST API roleassignments
    - Detects permission inheritance vs. unique permissions
    - Identifies "Everyone except external users" access
    
    Group Membership Validation:
    - SharePoint groups: Uses SharePoint REST API to check membership
    - Entra ID/M365 groups: Uses Graph API to check membership
    - Handles nested group scenarios
    
    Output Generation:
    - Creates detailed CSV report with access analysis
    - Generates comprehensive log file for troubleshooting
    - Provides color-coded console output based on debug setting
#>

#region Configuration
# App-Only Authentication Settings
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536" # This is your Entra App ID
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082" # This is certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3" # This is your Tenant ID

#Configurable Settings
$t = 'M365CPI13246019' # < - Your Tenant Name Here

# Single site and user to check
$siteUrl = 'https://m365cpi13246019.sharepoint.com/sites/SalesandMarketing'
$userToCheck = 'LisaT@M365CPI13246019.OnMicrosoft.com'

# Document identification - Use ONE of the following methods:
# METHOD 1: Direct URL (recommended for reliability) - Full SharePoint document URL
$documentUrl = 'https://m365cpi13246019.sharepoint.com/sites/SalesandMarketing/Shared%20Documents/Marketing/R%20and%20D%20Presentation.pdf'  # e.g., 'https://tenant.sharepoint.com/sites/sitename/Shared%20Documents/folder/document.pdf'

# Optional Feature Settings
$debug = $false  # Set to $true for detailed debug output, $false for minimal output
#endregion Configuration

# =================================================================================================
# END OF USER CONFIGURATION
# =================================================================================================


# =================================================================================================
# FUNCTION DEFINITIONS
# =================================================================================================

#region Authentication Functions
# Helper function to get access token using certificate-based authentication
Function Get-GraphAccessToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertThumbprint,
        [string]$Resource = "https://graph.microsoft.com/.default"  # Default to Graph API, can be changed to SharePoint
    )
    
    try {
        # Find the certificate
        $cert = Get-ChildItem -Path Cert:\CurrentUser\My\$CertThumbprint -ErrorAction SilentlyContinue
        if (-not $cert) {
            $cert = Get-ChildItem -Path Cert:\LocalMachine\My\$CertThumbprint -ErrorAction SilentlyContinue
        }
        
        if (-not $cert) {
            throw "Certificate with thumbprint $CertThumbprint not found"
        }
        
        # Create JWT token
        $now = [DateTime]::UtcNow
        $expiryDate = $now.AddMinutes(10)
        
        # Calculate Unix timestamps (seconds since 1970-01-01)
        $epoch = [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
        $nbf = [Math]::Floor(($now - $epoch).TotalSeconds)
        $exp = [Math]::Floor(($expiryDate - $epoch).TotalSeconds)
        
        # Create header
        $header = @{
            alg = "RS256"
            typ = "JWT"
            x5t = [Convert]::ToBase64String($cert.GetCertHash()) -replace '\+', '-' -replace '/', '_' -replace '='
        }
        
        # Create payload
        $payload = @{
            aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
            exp = $exp
            iss = $ClientId
            jti = [guid]::NewGuid().ToString()
            nbf = $nbf
            sub = $ClientId
        }
        
        # Encode header and payload
        $headerJson = $header | ConvertTo-Json -Compress
        $payloadJson = $payload | ConvertTo-Json -Compress
        
        $headerBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($headerJson)) -replace '\+', '-' -replace '/', '_' -replace '='
        $payloadBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadJson)) -replace '\+', '-' -replace '/', '_' -replace '='
        
        # Create signature
        $jwtContent = "$headerBase64.$payloadBase64"
        $jwtBytes = [System.Text.Encoding]::UTF8.GetBytes($jwtContent)
        
        $signature = $cert.PrivateKey.SignData($jwtBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $signatureBase64 = [Convert]::ToBase64String($signature) -replace '\+', '-' -replace '/', '_' -replace '='
        
        $jwt = "$jwtContent.$signatureBase64"
        
        # Request access token
        $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $body = @{
            client_id             = $ClientId
            client_assertion      = $jwt
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            scope                 = $Resource
            grant_type            = "client_credentials"
        }
        
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        throw "Failed to get access token: $($_.Exception.Message)"
    }
}
#endregion Authentication Functions

#region API Helper Functions
# Helper function to make Graph API calls
Function Invoke-CustomGraphRequest {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [string]$AccessToken,
        [object]$Body = $null
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    try {
        $params = @{
            Uri     = $Uri
            Method  = $Method
            Headers = $headers
        }
        
        if ($Body) {
            $params.Body = ($Body | ConvertTo-Json -Depth 10)
        }
        
        $response = Invoke-RestMethod @params
        return $response
    }
    catch {
        throw $_
    }
}

# Helper function to make SharePoint REST API calls
Function Invoke-SharePointRestRequest {
    param(
        [string]$SiteUrl,
        [string]$Endpoint,
        [string]$Method = "GET",
        [string]$AccessToken
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Accept"        = "application/json;odata=verbose"
        "Content-Type"  = "application/json;odata=verbose"
    }
    
    $uri = "$SiteUrl/_api/$Endpoint"
    
    try {
        # Use Invoke-WebRequest instead of Invoke-RestMethod to avoid PowerShell 7+ JSON parsing issues
        # with duplicate keys (Id vs ID) in SharePoint responses
        $webResponse = Invoke-WebRequest -Uri $uri -Method $Method -Headers $headers
        # Manually parse JSON with -AsHashtable to handle duplicate keys
        $response = $webResponse.Content | ConvertFrom-Json -AsHashtable -Depth 10
        return $response
    }
    catch {
        throw $_
    }
}

# Helper function to handle paginated Graph API results
Function Get-GraphAllPages {
    param(
        [string]$Uri,
        [string]$AccessToken
    )
    
    $allResults = @()
    $nextLink = $Uri
    
    do {
        $response = Invoke-CustomGraphRequest -Uri $nextLink -AccessToken $AccessToken
        
        if ($response.value) {
            $allResults += $response.value
        }
        else {
            $allResults += $response
        }
        
        $nextLink = $response.'@odata.nextLink'
    } while ($nextLink)
    
    return $allResults
}
#endregion API Helper Functions

#region Logging Functions
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
#endregion Logging Functions

#region Membership Check Functions
# Function to check if user is member of Entra ID group using Graph REST API
Function Test-EntraGroupMembership {
    param(
        [string] $UserObjectId,
        [string] $GroupId,
        [string] $GroupDisplayName,
        [string] $AccessToken
    )
    
    try {
        # Check if user is a member of the group
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members"
        $members = Get-GraphAllPages -Uri $uri -AccessToken $AccessToken
        
        $isMember = $members | Where-Object { $_.id -eq $UserObjectId }
        if ($isMember) {
            return $true
        }
        
        # Also check if user is an owner of the group
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/owners"
        $owners = Get-GraphAllPages -Uri $uri -AccessToken $AccessToken
        
        $isOwner = $owners | Where-Object { $_.id -eq $UserObjectId }
        if ($isOwner) {
            Write-DebugInfo "User is an OWNER of group '$GroupDisplayName'" -ForegroundColor Magenta
            return $true
        }
        
        return $false
    }
    catch {
        Write-DebugInfo "Could not check membership for group '$GroupDisplayName': $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $false
    }
}

# Function to check if user is member of SharePoint group using SharePoint REST API
Function Test-SharePointGroupMembership {
    param(
        [string] $SiteUrl,
        [string] $GroupId,
        [string] $GroupName,
        [string] $UserEmail
    )
    
    try {
        Write-DebugInfo "    Checking SharePoint group '$GroupName' (ID: $GroupId) for user $UserEmail" -ForegroundColor Yellow
        
        # Get group members from SharePoint REST API using SharePoint token
        $endpoint = "web/sitegroups/GetById($GroupId)/users"
        $result = Invoke-SharePointRestRequest -SiteUrl $SiteUrl -Endpoint $endpoint -AccessToken $script:sharePointToken
        
        if ($result.d.results) {
            $members = $result.d.results
            Write-DebugInfo "      Found $($members.Count) member(s) in group '$GroupName'" -ForegroundColor Gray
            
            foreach ($member in $members) {
                Write-DebugInfo "        Member: $($member.Title) ($($member.Email))" -ForegroundColor DarkGray
                if ($member.Email -eq $UserEmail -or $member.LoginName -like "*$UserEmail*") {
                    Write-DebugInfo "      ✓ FOUND user in SharePoint group '$GroupName'!" -ForegroundColor Green
                    return $true
                }
            }
            Write-DebugInfo "      User not found in group '$GroupName'" -ForegroundColor DarkGray
        }
        else {
            Write-DebugInfo "      Could not retrieve members for group '$GroupName'" -ForegroundColor DarkYellow
        }
        
        return $false
    }
    catch {
        Write-DebugInfo "      Error checking SharePoint group membership: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}
#endregion Membership Check Functions

# =================================================================================================
# CONNECTION AND INITIALIZATION
# =================================================================================================

#region Initialization
# Get Access Token for Microsoft Graph API using App-Only Authentication
try {
    Write-StatusMessage "Acquiring access token for Microsoft Graph API using certificate authentication..." -ForegroundColor Green
    $script:accessToken = Get-GraphAccessToken -TenantId $tenant -ClientId $appID -CertThumbprint $thumbprint -Resource "https://graph.microsoft.com/.default"
    Write-StatusMessage "Successfully acquired Graph API access token" -ForegroundColor Green
    
    # Also acquire SharePoint token for SharePoint REST API calls
    Write-StatusMessage "Acquiring access token for SharePoint REST API..." -ForegroundColor Green
    $sharePointResource = "https://$t.sharepoint.com/.default"
    $script:sharePointToken = Get-GraphAccessToken -TenantId $tenant -ClientId $appID -CertThumbprint $thumbprint -Resource $sharePointResource
    Write-StatusMessage "Successfully acquired SharePoint API access token" -ForegroundColor Green
    
    # Show debug mode status
    if ($debug) {
        Write-DebugInfo "DEBUG MODE: ENABLED - Detailed output will be shown" -ForegroundColor Cyan
    }
    else {
        Write-DebugInfo "DEBUG MODE: DISABLED - Minimal output will be shown" -ForegroundColor Gray
    }
}
catch {
    Write-StatusMessage "Failed to acquire access token: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

#Initialize Parameters - Do not change
$date = Get-Date -Format yyyy-MM-dd_HH-mm-ss

#OutPut and Log Files
$outputfile = "$env:TEMP\" + 'CheckAccess_' + $date + "output.csv"
$log = "$env:TEMP\" + 'CheckAccess_' + $date + '_' + "logfile.log"

# Get the specific site using Graph REST API
$uri = "https://graph.microsoft.com/v1.0/sites?`$select=id,displayName,webUrl,siteCollection,createdDateTime"
$allSites = Get-GraphAllPages -Uri $uri -AccessToken $script:accessToken
$site = $allSites | Where-Object { $_.webUrl -eq $siteUrl }

if (-not $site) {
    throw "Site not found: $siteUrl"
}

if (-not $site) {
    Write-StatusMessage "Could not find site: $siteUrl" -ForegroundColor Red
    exit
}

Write-StatusMessage "Found site: $($site.displayName) ($($site.webUrl))" -ForegroundColor Green
#endregion Initialization

#region Document Retrieval
# Get the document - either by direct URL or by searching
if ($documentUrl -and $documentUrl -ne '') {
    Write-StatusMessage "Retrieving document by URL: $documentUrl" -ForegroundColor Cyan
    
    # Parse the URL to extract drive and file information
    # URL format: https://tenant.sharepoint.com/sites/sitename/Library/path/to/file.ext
    $document = $null
    
    # Try to get the item by URL path
    # Extract the path after the site URL
    $urlPath = $documentUrl -replace [regex]::Escape($site.webUrl), ''
    $urlPath = $urlPath.TrimStart('/')
    
    # URL decode the path
    $urlPath = [System.Web.HttpUtility]::UrlDecode($urlPath)
    
    # Try to determine the drive and item path
    # Common patterns: /Shared Documents/, /Documents/, /LibraryName/
    if ($urlPath -match '^([^/]+)/(.+)$') {
        $libraryName = $matches[1]
        $itemPath = $matches[2]
            
        Write-DebugInfo "  Library: $libraryName, Item path: $itemPath"
        
        # STRATEGY: Use SharePoint REST API directly to get document by server-relative path
        # This is more reliable than Graph API for URL-based retrieval
        
        # Build server-relative path
        $sitePath = ($site.webUrl -replace '^https?://[^/]+', '')
        
        # Map library name to SharePoint internal name
        $libraryInternalName = $libraryName
        if ($libraryName -eq 'Documents') {
            $libraryInternalName = 'Shared Documents'
        }
        
        $serverRelativePath = "$sitePath/$libraryInternalName/$itemPath"
        Write-DebugInfo "  Server-relative path for retrieval: $serverRelativePath" -ForegroundColor Cyan
        
        try {
            # Use SharePoint REST API to get file by server-relative path
            $endpoint = "web/GetFileByServerRelativePath(decodedurl='$serverRelativePath')/ListItemAllFields"
            $spListItemResponse = Invoke-SharePointRestRequest `
                -SiteUrl $site.webUrl `
                -Endpoint $endpoint `
                -AccessToken $script:sharePointToken
            
            if ($spListItemResponse -and $spListItemResponse.d) {
                $listItemId = $spListItemResponse.d.ID -as [string]
                if (-not $listItemId) {
                    $listItemId = $spListItemResponse.d.Item('ID') -as [string]
                }
                
                Write-DebugInfo "  ✓ Found document via SharePoint REST API (Item ID: $listItemId)" -ForegroundColor Green
                
                # Get author information from SharePoint
                $ownerEmail = $null
                $ownerName = $null
                try {
                    $authorEndpoint = "web/GetFileByServerRelativePath(decodedurl='$serverRelativePath')/ListItemAllFields/Author"
                    $authorResponse = Invoke-SharePointRestRequest `
                        -SiteUrl $site.webUrl `
                        -Endpoint $authorEndpoint `
                        -AccessToken $script:sharePointToken
                    
                    if ($authorResponse.d) {
                        if ($authorResponse.d.Email) { $ownerEmail = $authorResponse.d.Email }
                        if ($authorResponse.d.Title) { $ownerName = $authorResponse.d.Title }
                        
                        # If we have a name but no email, try to look up the user in Entra ID
                        if ($ownerName -and -not $ownerEmail) {
                            Write-DebugInfo "  No email in SharePoint, looking up user in Entra ID by display name..." -ForegroundColor Yellow
                            try {
                                $userSearchUri = "https://graph.microsoft.com/v1.0/users?`$filter=displayName eq '$($ownerName.Replace("'", "''"))'&`$select=mail,userPrincipalName,displayName"
                                $userSearchResult = Invoke-CustomGraphRequest -Uri $userSearchUri -AccessToken $script:accessToken
                                
                                if ($userSearchResult.value -and $userSearchResult.value.Count -gt 0) {
                                    $foundUser = $userSearchResult.value[0]
                                    $ownerEmail = if ($foundUser.mail) { $foundUser.mail } else { $foundUser.userPrincipalName }
                                    Write-DebugInfo "  ✓ Found user email in Entra ID: $ownerEmail" -ForegroundColor Green
                                }
                                else {
                                    Write-DebugInfo "  Could not find user in Entra ID by display name" -ForegroundColor DarkYellow
                                }
                            }
                            catch {
                                Write-DebugInfo "  Error looking up user in Entra ID: $($_.Exception.Message)" -ForegroundColor DarkYellow
                            }
                        }
                        
                        $emailDisplay = if ($ownerEmail) { $ownerEmail } else { "no email" }
                        Write-DebugInfo "  ✓ Retrieved author: $ownerName ($emailDisplay)" -ForegroundColor Green
                    }
                }
                catch {
                    Write-DebugInfo "  Could not retrieve author information: $($_.Exception.Message)" -ForegroundColor DarkYellow
                }
                
                # Now get drives to map back to Graph API structure
                $drivesUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives"
                $drives = Get-GraphAllPages -Uri $drivesUri -AccessToken $script:accessToken
                
                $targetDrive = $drives | Where-Object { 
                    $_.name -eq $libraryName -or 
                    $_.name -eq ($libraryName -replace ' ', '') -or
                    ($libraryName -eq 'Shared Documents' -and $_.name -eq 'Documents') -or
                    ($libraryName -eq 'Documents' -and $_.name -eq 'Shared Documents')
                }
                
                if ($targetDrive -and $listItemId) {
                    # Get document via Graph API using the SharePoint list item ID
                    try {
                        $itemUri = "https://graph.microsoft.com/v1.0/drives/$($targetDrive.id)/items/$listItemId" + "?`$select=*,sharepointIds&`$expand=createdBy"
                        $document = Invoke-CustomGraphRequest -Uri $itemUri -AccessToken $script:accessToken
                        Write-DebugInfo "  ✓ Retrieved full document details from Graph API" -ForegroundColor Green
                        
                        # If Graph API didn't return createdBy, use SharePoint author info
                        if (-not $document.createdBy -or -not $document.createdBy.user -or -not $document.createdBy.user.email) {
                            if ($ownerEmail -ne "N/A" -or $ownerName -ne "N/A") {
                                $document | Add-Member -MemberType NoteProperty -Name "createdBy" -Value @{
                                    user = @{
                                        email       = $ownerEmail
                                        displayName = $ownerName
                                    }
                                } -Force
                                Write-DebugInfo "  ✓ Added author info from SharePoint to Graph document" -ForegroundColor Green
                            }
                        }
                    }
                    catch {
                        Write-DebugInfo "  Could not retrieve document from Graph API using list item ID: $($_.Exception.Message)" -ForegroundColor DarkYellow
                        
                        # Create a minimal document object with required properties from SharePoint data
                        $fileName = if ($itemPath -match '([^/]+)$') { $matches[1] } else { "Unknown" }
                        
                        # Build parent path from item path (remove filename)
                        $folderPath = if ($itemPath -match '^(.+)/[^/]+$') { "/" + $matches[1] } else { "" }
                        
                        # Use owner info already retrieved from SharePoint REST API
                        
                        $document = @{
                            id                    = $listItemId
                            name                  = $fileName
                            webUrl                = $documentUrl
                            sharepointIds         = @{
                                listItemId = $listItemId
                                siteId     = $site.id
                            }
                            parentReference       = @{
                                driveId = $targetDrive.id
                                siteId  = $site.id
                                # Add path info for validation to work
                                path    = "/drives/$($targetDrive.id)/root:$folderPath"
                            }
                            # Add owner info from SharePoint
                            createdBy             = @{
                                user = @{
                                    email       = $ownerEmail
                                    displayName = $ownerName
                                }
                            }
                            # Store the drive name for validation
                            _spLibraryName        = $libraryName
                            # Store the server-relative path that we already validated
                            _spServerRelativePath = $serverRelativePath
                        }
                        $emailDisplay = if ($ownerEmail) { $ownerEmail } else { "no email" }
                        $nameDisplay = if ($ownerName) { $ownerName } else { "no name" }
                        Write-DebugInfo "  Created minimal document object from SharePoint data (Owner: $nameDisplay / $emailDisplay)" -ForegroundColor Cyan
                    }
                }
            }
        }
        catch {
            Write-DebugInfo "  Could not retrieve document via SharePoint REST API: $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }
    
    if (-not $document) {
        Write-StatusMessage "ERROR: Could not retrieve document from URL." -ForegroundColor Red
        
        # Extract filename from URL for error message
        if ($urlPath -match '([^/]+)$') {
            $documentName = [System.Web.HttpUtility]::UrlDecode($matches[1])
            Write-DebugInfo "  Document name from URL: $documentName" -ForegroundColor Cyan
        }
        
        Write-StatusMessage "Search fallback removed to ensure accuracy when multiple files with same name exist." -ForegroundColor Red
        Write-StatusMessage "Please verify the document URL is correct and the document exists." -ForegroundColor Red
        continue
    }
}

# Document name search removed - URL-based retrieval required for accuracy
# to avoid finding wrong files when multiple files share the same name

if (-not $document) {
    Write-StatusMessage "Document not found: $documentName" -ForegroundColor Red
    Write-LogEntry -LogName:$Log -LogEntryText "Document not found: $documentName in site $($site.webUrl)"
    continue
}

Write-StatusMessage "Found document: $($document.name) (ID: $($document.id))" -ForegroundColor Green
Write-StatusMessage "Document path: $($document.webUrl)" -ForegroundColor Gray
if ($document.eTag -and $document.cTag) {
    Write-DebugInfo "Document properties: eTag=$($document.eTag), cTag=$($document.cTag)" -ForegroundColor DarkGray
}

#endregion Document Retrieval

# =================================================================================================
# MAIN PROCESSING
# =================================================================================================

#region Permission Checking
# Initialize site-specific output array
$siteOutput = @()

Write-StatusMessage "Processing document: $($document.name) in site: $($site.displayName)" -ForegroundColor Yellow
Write-LogEntry -LogName:$Log -LogEntryText "Starting processing for site: $($site.displayName) ($($site.webUrl))"

Write-StatusMessage "Checking access for user: $userToCheck on document: $($document.name)" -ForegroundColor Cyan

# Check user permissions on document
Write-DebugInfo "Attempting to check '$userToCheck' permissions on DOCUMENT '$($document.name)' in SITE '$($site.webUrl)'" -ForegroundColor Green
Write-LogEntry -LogName:$Log -LogEntryText "Attempting to check '$userToCheck' permissions on DOCUMENT '$($document.name)' in SITE '$($site.webUrl)'"

try {
    $userFound = $false
    $accessType = ""
    $groupMemberships = @()
    $eeeuFound = $false
    $eeeuRoles = @()
    $hasMeaningfulDirectAccess = $false
    
    # Get the user object from Graph REST API
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$userToCheck'"
    $result = Invoke-CustomGraphRequest -Uri $uri -AccessToken $script:accessToken
    $userObject = $result.value | Select-Object -First 1
    
    if (-not $userObject) {
        Write-StatusMessage "Could not find user object for '$userToCheck'" -ForegroundColor Red
        Write-LogEntry -LogName:$Log -LogEntryText "Could not find user object for '$userToCheck'"
    }
    else {
        #region Direct Permission Checks
        # First, check if the document has unique permissions (broken inheritance)
        $hasUniquePermissions = $false
        try {
            $docWebUrl = $document.webUrl
            if ($docWebUrl -match "$([regex]::Escape($site.webUrl))(.+)$") {
                $serverRelativeUrl = $matches[1]
                $serverRelativeUrl = $serverRelativeUrl -replace '\?.*$', ''
                $serverRelativeUrl = [System.Web.HttpUtility]::UrlDecode($serverRelativeUrl)
                
                if ($site.webUrl -match 'https://[^/]+(/sites/[^/]+)') {
                    $sitePath = $matches[1]
                    $fullServerRelativeUrl = "$sitePath$serverRelativeUrl"
                }
                else {
                    $fullServerRelativeUrl = $serverRelativeUrl
                }
                
                $encodedUrl = [System.Web.HttpUtility]::UrlEncode($fullServerRelativeUrl).Replace('+', '%20')
                $endpoint = "web/GetFileByServerRelativeUrl('$encodedUrl')/ListItemAllFields?`$select=HasUniqueRoleAssignments"
                $itemData = Invoke-SharePointRestRequest -SiteUrl $site.webUrl -Endpoint $endpoint -AccessToken $script:sharePointToken
                
                if ($itemData.d.HasUniqueRoleAssignments -ne $null) {
                    $hasUniquePermissions = $itemData.d.HasUniqueRoleAssignments
                    if ($hasUniquePermissions) {
                        Write-DebugInfo "✓ Document has UNIQUE permissions (inheritance is broken)" -ForegroundColor Magenta
                    }
                    else {
                        Write-DebugInfo "Document inherits permissions from parent" -ForegroundColor Cyan
                    }
                }
            }
        }
        catch {
            Write-DebugInfo "Could not check inheritance status: $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
        
        # Check 1: Direct user permissions on the document using Graph REST API
        Write-DebugInfo "Checking direct document permissions..." -ForegroundColor Yellow
        $uri = "https://graph.microsoft.com/v1.0/drives/$($document.parentReference.driveId)/items/$($document.id)/permissions"
        $documentPermissions = Get-GraphAllPages -Uri $uri -AccessToken $script:accessToken
        
        Write-DebugInfo "Found $($documentPermissions.Count) permission(s) on document (from Graph API)" -ForegroundColor DarkYellow
        
        # IMPORTANT: If document has unique permissions, Graph API sometimes returns stale/inherited permissions
        # We need to validate against SharePoint REST API roleassignments for accurate results
        $actualGroupsOnDocument = @()
        if ($hasUniquePermissions) {
            Write-DebugInfo "Document has unique permissions - validating groups via SharePoint REST API..." -ForegroundColor Yellow
            try {
                # Get list item ID and library title
                $listItemId = $null
                $libraryTitle = $null
                
                if ($document.parentReference.driveId) {
                    $driveId = $document.parentReference.driveId
                    $drivesUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives"
                    $drives = Get-GraphAllPages -Uri $drivesUri -AccessToken $script:accessToken
                    $targetDrive = $drives | Where-Object { $_.id -eq $driveId }
                    if ($targetDrive) {
                        $libraryTitle = $targetDrive.name
                    }
                }
                
                if ($document._spServerRelativePath -and $document.sharepointIds.listItemId) {
                    $listItemId = $document.sharepointIds.listItemId
                }
                
                if ($listItemId -and $libraryTitle) {
                    $endpoint = "web/lists/GetByTitle('$libraryTitle')/items($listItemId)/roleassignments?`$expand=Member,RoleDefinitionBindings"
                    $roleAssignments = Invoke-SharePointRestRequest -SiteUrl $site.webUrl -Endpoint $endpoint -AccessToken $script:sharePointToken
                    
                    if ($roleAssignments.d.results) {
                        Write-DebugInfo "  ✓ Found $($roleAssignments.d.results.Count) actual role assignment(s) on document" -ForegroundColor Green
                        
                        # Build list of actual groups that have permissions on this document
                        foreach ($ra in $roleAssignments.d.results) {
                            $memberTitle = $ra.Member.Title
                            $memberLoginName = $ra.Member.LoginName
                            $memberId = $ra.Member.Id
                            $memberPrincipalType = $ra.Member.PrincipalType
                            
                            # Extract groups - PrincipalType: 1=User, 4=SharePointGroup, 8=SecurityGroup (Entra)
                            # Also filter by loginName patterns
                            $isUser = ($memberLoginName -like "*membership*" -and $memberLoginName -notlike "*spo-grid-all-users*")
                            $isGroup = ($memberPrincipalType -eq 4 -or $memberPrincipalType -eq 8 -or 
                                $memberLoginName -like "c:0*.c|*" -or 
                                $memberLoginName -like "c:0t.c|*" -or
                                $memberLoginName -notlike "*membership*")
                            
                            if ($isGroup -and -not $isUser) {
                                # This is a group (SharePoint or Entra/M365)
                                $roles = $ra.RoleDefinitionBindings.results | ForEach-Object { $_.Name }
                                $actualGroupsOnDocument += @{
                                    Name          = $memberTitle
                                    Id            = $memberId
                                    LoginName     = $memberLoginName
                                    Roles         = $roles
                                    PrincipalType = $memberPrincipalType
                                }
                                Write-DebugInfo "    Actual group on document: $memberTitle (ID: $memberId, Type: $memberPrincipalType, Roles: $($roles -join ','))" -ForegroundColor Cyan
                            }
                        }
                    }
                }
            }
            catch {
                Write-DebugInfo "  Could not validate groups via roleassignments: $($_.Exception.Message)" -ForegroundColor DarkYellow
                # If validation fails, we'll have to trust Graph API (but warn about it)
                Write-DebugInfo "  ⚠ WARNING: Cannot validate permissions - Graph API may return stale data" -ForegroundColor Yellow
            }
        }
        
        # Check each permission entry
        foreach ($permission in $documentPermissions) {
            Write-DebugInfo "  Permission ID: $($permission.id), Roles: $($permission.roles -join ', ')" -ForegroundColor Gray
            
            # Dump the full permission object for debugging
            Write-DebugInfo "    Full permission object: $($permission | ConvertTo-Json -Depth 5 -Compress)" -ForegroundColor DarkGray
            
            # Check if this permission applies to our user directly
            if ($permission.grantedToIdentitiesV2) {
                foreach ($identity in $permission.grantedToIdentitiesV2) {
                    Write-DebugInfo "    Checking grantedToIdentitiesV2: $($identity | ConvertTo-Json -Depth 3 -Compress)" -ForegroundColor DarkGray
                    if ($identity.user.email -eq $userToCheck -or $identity.user.displayName -eq $userObject.displayName -or $identity.user.id -eq $userObject.id) {
                        $userFound = $true
                        # Track if this is meaningful access (not just Limited Access)
                        if ($permission.roles -notcontains "read" -and $permission.roles -notcontains "write" -and $permission.roles -notcontains "owner") {
                            $accessType += "; Limited Access"
                        }
                        else {
                            $accessType += "; Direct Document Access ($($permission.roles -join ', '))"
                            $hasMeaningfulDirectAccess = $true
                        }
                        Write-DebugInfo "✓ Found $userToCheck with DIRECT access on document: $($permission.roles -join ', ')" -ForegroundColor Green
                    }
                }
            }
            
            # Check if this permission applies to our user via grantedTo (older API)
            if ($permission.grantedTo) {
                Write-DebugInfo "    Checking grantedTo: $($permission.grantedTo | ConvertTo-Json -Depth 3 -Compress)" -ForegroundColor DarkGray
                if ($permission.grantedTo.user.email -eq $userToCheck -or $permission.grantedTo.user.displayName -eq $userObject.displayName -or $permission.grantedTo.user.id -eq $userObject.id) {
                    $userFound = $true
                    # Track if this is meaningful access (not just Limited Access)
                    if ($permission.roles -notcontains "read" -and $permission.roles -notcontains "write" -and $permission.roles -notcontains "owner") {
                        $accessType += "; Limited Access"
                    }
                    else {
                        $accessType += "; Direct Document Access ($($permission.roles -join ', '))"
                        $hasMeaningfulDirectAccess = $true
                    }
                    Write-DebugInfo "✓ Found $userToCheck with DIRECT access (grantedTo) on document: $($permission.roles -join ', ')" -ForegroundColor Green
                }
            }
            
            # Check if permission is granted via grantedToV2
            if ($permission.grantedToV2) {
                Write-DebugInfo "    Checking grantedToV2: $($permission.grantedToV2 | ConvertTo-Json -Depth 3 -Compress)" -ForegroundColor DarkGray
                
                # Check for "Everyone except external users" permissions in siteUser FIRST
                # This prevents EEEU permissions from being misidentified as direct user access
                $isEEEUPermission = $false
                if ($permission.grantedToV2.siteUser) {
                    $loginName = $permission.grantedToV2.siteUser.loginName
                    $displayName = $permission.grantedToV2.siteUser.displayName
                    
                    if ($loginName -like "*spo-grid-all-users*" -or $displayName -eq "Everyone except external users") {
                        Write-DebugInfo "    Found 'Everyone except external users' permission with roles: $($permission.roles -join ', ')" -ForegroundColor Cyan
                        $eeeuFound = $true
                        $eeeuRoles += $permission.roles
                        $isEEEUPermission = $true
                    }
                }
                
                # Only check direct user access if this is NOT an EEEU permission
                if (-not $isEEEUPermission -and $permission.grantedToV2.user) {
                    if ($permission.grantedToV2.user.email -eq $userToCheck -or $permission.grantedToV2.user.displayName -eq $userObject.displayName -or $permission.grantedToV2.user.id -eq $userObject.id) {
                        $userFound = $true
                        # Track if this is meaningful access (not just Limited Access)
                        if ($permission.roles -notcontains "read" -and $permission.roles -notcontains "write" -and $permission.roles -notcontains "owner") {
                            $accessType += "; Limited Access"
                        }
                        else {
                            $accessType += "; Direct Document Access ($($permission.roles -join ', '))"
                            $hasMeaningfulDirectAccess = $true
                        }
                        Write-DebugInfo "✓ Found $userToCheck with DIRECT access (grantedToV2) on document: $($permission.roles -join ', ')" -ForegroundColor Green
                    }
                }
                
                # Check SharePoint group access
                if ($permission.grantedToV2.siteGroup) {
                    $spGroupId = $permission.grantedToV2.siteGroup.id
                    $spGroupName = $permission.grantedToV2.siteGroup.displayName
                    Write-DebugInfo "    Found SharePoint group in Graph API permissions: $spGroupName (ID: $spGroupId)" -ForegroundColor Yellow
                    
                    # If document has unique permissions and we validated via roleassignments, 
                    # verify this group is actually on the document (not stale Graph API data)
                    $groupIsActuallyOnDocument = $true
                    if ($hasUniquePermissions -and $actualGroupsOnDocument.Count -gt 0) {
                        # Check if this group is in our validated list
                        $groupIsActuallyOnDocument = $actualGroupsOnDocument | Where-Object { 
                            $_.Id -eq $spGroupId -or $_.Name -eq $spGroupName 
                        }
                        
                        if (-not $groupIsActuallyOnDocument) {
                            Write-DebugInfo "    ⚠ SKIPPING: Group not found in actual roleassignments - Graph API returned stale data" -ForegroundColor Red
                        }
                    }
                    
                    if ($groupIsActuallyOnDocument) {
                        # This group has permissions ON THIS DOCUMENT
                        # Now check if user is member of this SharePoint group
                        $isMemberOfSPGroup = Test-SharePointGroupMembership -SiteUrl $site.webUrl -GroupId $spGroupId -GroupName $spGroupName -UserEmail $userToCheck
                        if ($isMemberOfSPGroup) {
                            $userFound = $true
                            $accessType += "; Via SharePoint Group: $spGroupName ($($permission.roles -join ', '))"
                            $groupMemberships += "SharePoint Group: $spGroupName"
                            Write-DebugInfo "    ✓ User is member of SharePoint group '$spGroupName' which has access to this document!" -ForegroundColor Green
                        }
                        else {
                            Write-DebugInfo "    User is NOT a member of group '$spGroupName'" -ForegroundColor DarkGray
                        }
                    }
                }
                
                # Check M365/Entra ID group access
                if ($permission.grantedToV2.group) {
                    $m365GroupId = $permission.grantedToV2.group.id
                    $m365GroupName = $permission.grantedToV2.group.displayName
                    
                    # Check for "Everyone except external users" - don't treat it as a normal M365 group
                    if ($m365GroupName -eq "Everyone except external users") {
                        Write-DebugInfo "    Found 'Everyone except external users' permission (as group) with roles: $($permission.roles -join ', ')" -ForegroundColor Cyan
                        $eeeuFound = $true
                        $eeeuRoles += $permission.roles
                    }
                    else {
                        Write-DebugInfo "    Found M365/Entra ID group in Graph API permissions: $m365GroupName (ID: $m365GroupId)" -ForegroundColor Yellow
                        
                        # If document has unique permissions and we validated via roleassignments,
                        # verify this group is actually on the document (not stale Graph API data)
                        $groupIsActuallyOnDocument = $true
                        if ($hasUniquePermissions -and $actualGroupsOnDocument.Count -gt 0) {
                            # M365 groups in roleassignments have different IDs, match by name
                            $groupIsActuallyOnDocument = $actualGroupsOnDocument | Where-Object { 
                                $_.Name -eq $m365GroupName 
                            }
                            
                            if (-not $groupIsActuallyOnDocument) {
                                Write-DebugInfo "    ⚠ SKIPPING: Group not found in actual roleassignments - Graph API returned stale data" -ForegroundColor Red
                            }
                        }
                        
                        if ($groupIsActuallyOnDocument) {
                            # This group has permissions ON THIS DOCUMENT
                            # Now check if user is member of this M365/Entra ID group
                            $isMemberOfM365Group = Test-EntraGroupMembership -UserObjectId $userObject.id -GroupId $m365GroupId -GroupDisplayName $m365GroupName -AccessToken $script:accessToken
                            if ($isMemberOfM365Group) {
                                $userFound = $true
                                $accessType += "; Via M365 Group: $m365GroupName ($($permission.roles -join ', '))"
                                $groupMemberships += "M365 Group: $m365GroupName"
                                Write-DebugInfo "    ✓ User is member of M365 group '$m365GroupName' which has access to this document!" -ForegroundColor Green
                            }
                            else {
                                Write-DebugInfo "    User is NOT a member of group '$m365GroupName'" -ForegroundColor DarkGray
                            }
                        }
                    }
                }
            }
            
            # Check if permission is granted via a group
            if ($permission.grantedToIdentitiesV2) {
                foreach ($identity in $permission.grantedToIdentitiesV2) {
                    if ($identity.siteGroup) {
                        $groupId = $identity.siteGroup.id
                        $groupName = $identity.siteGroup.displayName
                        Write-DebugInfo "  Checking if user is in SharePoint group: $groupName" -ForegroundColor Yellow
                        
                        # Note: Graph API doesn't support checking SharePoint group membership directly
                        # These are inherited permissions, so we'll rely on site-level checks
                        Write-DebugInfo "    SharePoint group membership will be checked via site permissions" -ForegroundColor DarkGray
                    }
                }
            }
            
            # Check if permission is inherited from parent (site/library)
            if ($permission.inheritedFrom) {
                Write-DebugInfo "  Permission is inherited from: $($permission.inheritedFrom.name)" -ForegroundColor DarkCyan
            }
            
            # Check sharing links
            if ($permission.link) {
                Write-DebugInfo "  Permission has sharing link: scope=$($permission.link.scope), type=$($permission.link.type)" -ForegroundColor DarkCyan
                if ($permission.link.scope -eq "organization" -or $permission.link.scope -eq "users") {
                    Write-DebugInfo "    This is an organization/user sharing link" -ForegroundColor DarkCyan
                }
            }
        }
        
        # Only check parent library permissions if document INHERITS from parent
        # Skip parent check if document has unique permissions (broken inheritance)
        if (-not $hasUniquePermissions) {
            Write-DebugInfo "Document inherits permissions - checking parent library permissions..." -ForegroundColor Yellow
            try {
                # Get the library (drive) that contains this document
                $driveId = $document.parentReference.driveId
                $uri = "https://graph.microsoft.com/v1.0/drives/$driveId/root/permissions"
                $libraryPermissions = Get-GraphAllPages -Uri $uri -AccessToken $script:accessToken
                
                Write-DebugInfo "Found $($libraryPermissions.Count) permission(s) on parent library" -ForegroundColor DarkYellow
                
                # Check library permissions for user access
                foreach ($libPerm in $libraryPermissions) {
                    if ($libPerm.grantedToV2.siteGroup) {
                        $spGroupId = $libPerm.grantedToV2.siteGroup.id
                        $spGroupName = $libPerm.grantedToV2.siteGroup.displayName
                        Write-DebugInfo "  Library has SharePoint group: $spGroupName" -ForegroundColor Yellow
                        
                        # Try to get members of this SharePoint group
                        # Note: This requires SharePoint REST API, not Graph API
                        # For now, we'll note these groups exist
                    }
                }
            }
            catch {
                Write-DebugInfo "Could not check parent library permissions: $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }
        else {
            Write-DebugInfo "Skipping parent library check - document has unique permissions" -ForegroundColor Magenta
        }
        
        # Check for groups in roleassignments that Graph API didn't return
        # This handles cases where Graph API is incomplete/stale
        if ($hasUniquePermissions -and $actualGroupsOnDocument.Count -gt 0) {
            Write-DebugInfo "Checking for groups in roleassignments that Graph API missed..." -ForegroundColor Yellow
            
            foreach ($actualGroup in $actualGroupsOnDocument) {
                # Check if this group was in Graph API permissions
                $groupFoundInGraphAPI = $false
                foreach ($permission in $documentPermissions) {
                    if ($permission.grantedToV2.siteGroup.displayName -eq $actualGroup.Name -or
                        $permission.grantedToV2.siteGroup.id -eq $actualGroup.Id -or
                        $permission.grantedToV2.group.displayName -eq $actualGroup.Name) {
                        $groupFoundInGraphAPI = $true
                        break
                    }
                }
                
                if (-not $groupFoundInGraphAPI) {
                    Write-DebugInfo "  Group '$($actualGroup.Name)' found in roleassignments but NOT in Graph API - checking membership..." -ForegroundColor Yellow
                    
                    # Try checking as SharePoint group first
                    $isMember = $false
                    $foundAsSharePointGroup = $false
                    
                    try {
                        $isMember = Test-SharePointGroupMembership -SiteUrl $site.webUrl -GroupId $actualGroup.Id -GroupName $actualGroup.Name -UserEmail $userToCheck
                        $foundAsSharePointGroup = $true
                    }
                    catch {
                        Write-DebugInfo "    Not a SharePoint group or error checking: $($_.Exception.Message)" -ForegroundColor DarkGray
                    }
                    
                    # If SharePoint check failed or user not found, try as Entra group
                    if (-not $foundAsSharePointGroup -or (-not $isMember)) {
                        Write-DebugInfo "    Trying as Entra ID group..." -ForegroundColor Yellow
                        try {
                            $groupSearchUri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$($actualGroup.Name.Replace("'", "''"))'&`$select=id,displayName"
                            $groupSearchResult = Invoke-CustomGraphRequest -Uri $groupSearchUri -AccessToken $script:accessToken
                            
                            if ($groupSearchResult.value -and $groupSearchResult.value.Count -gt 0) {
                                $entraGroup = $groupSearchResult.value[0]
                                $isMember = Test-EntraGroupMembership -UserObjectId $userObject.id -GroupId $entraGroup.id -GroupDisplayName $entraGroup.displayName -AccessToken $script:accessToken
                                $foundAsSharePointGroup = $false  # Mark that we found it as Entra group
                            }
                        }
                        catch {
                            Write-DebugInfo "    Error searching Entra ID: $($_.Exception.Message)" -ForegroundColor DarkYellow
                        }
                    }
                    
                    if ($isMember) {
                        $userFound = $true
                        $roleDisplay = if ($actualGroup.Roles -contains "Edit" -or $actualGroup.Roles -contains "Contribute") { "write" } 
                        elseif ($actualGroup.Roles -contains "Full Control") { "owner" } 
                        elseif ($actualGroup.Roles -contains "Read") { "read" } 
                        else { $actualGroup.Roles -join ',' }
                        $groupType = if ($foundAsSharePointGroup) { "SharePoint" } else { "Entra" }
                        $accessType += "; Via $groupType Group: $($actualGroup.Name) ($roleDisplay)"
                        Write-DebugInfo "    ✓ User is member of $groupType group '$($actualGroup.Name)' (found via roleassignments)" -ForegroundColor Green
                    }
                }
            }
        }
        #endregion Direct Permission Checks
        
        #region Effective Permission Validation
        # Fallback: Use SharePoint REST API to check effective permissions if Graph API found nothing
        if (-not $userFound -and -not $eeeuFound) {
            Write-DebugInfo "Graph API found no access. Checking effective permissions via SharePoint REST API..." -ForegroundColor Yellow
            try {
                # Get the document's server-relative URL from webUrl
                $docWebUrl = $document.webUrl
                if ($docWebUrl -match "$([regex]::Escape($site.webUrl))(.+)$") {
                    $serverRelativeUrl = $matches[1]
                    # Remove any query parameters
                    $serverRelativeUrl = $serverRelativeUrl -replace '\?.*$', ''
                    # URL decode the path
                    $serverRelativeUrl = [System.Web.HttpUtility]::UrlDecode($serverRelativeUrl)
                    
                    # Need full server-relative URL including site path
                    # Extract site path from site URL
                    if ($site.webUrl -match 'https://[^/]+(/sites/[^/]+)') {
                        $sitePath = $matches[1]
                        $fullServerRelativeUrl = "$sitePath$serverRelativeUrl"
                    }
                    else {
                        $fullServerRelativeUrl = $serverRelativeUrl
                    }
                    
                    Write-DebugInfo "  Full server-relative URL: $fullServerRelativeUrl" -ForegroundColor DarkGray
                    
                    # Use SharePoint REST API to get user's effective permissions
                    # First, we need to get the user's login name in SharePoint format
                    $userLoginName = "i:0#.f|membership|$($userToCheck.ToLower())"
                    $encodedLoginName = [System.Web.HttpUtility]::UrlEncode($userLoginName)
                    $encodedUrl = [System.Web.HttpUtility]::UrlEncode($fullServerRelativeUrl).Replace('+', '%20')
                    $endpoint = "web/GetFileByServerRelativeUrl('$encodedUrl')/ListItemAllFields/GetUserEffectivePermissions(@user)?@user='$encodedLoginName'"
                    
                    Write-DebugInfo "  Calling: $endpoint" -ForegroundColor DarkGray
                    
                    try {
                        $effectivePerms = Invoke-SharePointRestRequest -SiteUrl $site.webUrl -Endpoint $endpoint -AccessToken $script:sharePointToken
                        if ($effectivePerms.d) {
                            $high = $effectivePerms.d.GetUserEffectivePermissions.High
                            $low = $effectivePerms.d.GetUserEffectivePermissions.Low
                            
                            Write-DebugInfo "  Permission bits - High: $high, Low: $low" -ForegroundColor DarkGray
                            
                            # Check if user has meaningful permissions (not just Limited Access)
                            # Permissions: ViewListItems=0x1, AddListItems=0x2, EditListItems=0x4, DeleteListItems=0x8
                            # OpenItems=0x10, ViewVersions=0x20, DeleteVersions=0x40, Open=0x10000
                            $hasEdit = ($low -band 0x4) -ne 0
                            $hasView = ($low -band 0x1) -ne 0
                            $hasOpen = ($low -band 0x10000) -ne 0
                            
                            if ($hasEdit -or ($hasView -and $hasOpen)) {
                                $userFound = $true
                                $permLevel = if ($hasEdit) { "edit" } else { "read" }
                                $accessType = "Via Everyone Except External Users ($permLevel)"
                                $eeeuFound = $true
                                $eeeuRoles = @($permLevel)
                                Write-DebugInfo "  ✓ User has effective $permLevel permission (detected via SharePoint REST API)" -ForegroundColor Green
                            }
                            else {
                                Write-DebugInfo "  User has only Limited Access (no meaningful permissions)" -ForegroundColor DarkYellow
                            }
                        }
                    }
                    catch {
                        Write-DebugInfo "  Could not check effective permissions: $($_.Exception.Message)" -ForegroundColor DarkYellow
                    }
                }
            }
            catch {
                Write-DebugInfo "Could not parse document URL for SharePoint REST API check: $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }
        
        # Check if user qualifies for EEEU access (internal user)
        $userQualifiesForEEEU = $false
        $isInternalUser = ($userObject.userType -eq "Member" -or $userToCheck -like "*@*.onmicrosoft.com")
        
        # Run validation if:
        # 1. EEEU permission was found in Graph API, OR
        # 2. User has access but we need to verify if it's truly direct or actually via EEEU
        #    (Graph API sometimes incorrectly expands EEEU permissions as direct user access)
        if ($eeeuFound -or ($userFound -and $isInternalUser)) {
            # Check if user is internal (not external)
            if ($isInternalUser) {
                if ($eeeuFound) {
                    $userQualifiesForEEEU = $true
                    Write-DebugInfo "User qualifies for 'Everyone except external users' access (roles: $($eeeuRoles -join ', '))" -ForegroundColor Cyan
                }
                else {
                    Write-DebugInfo "No EEEU permission detected by Graph API, but validating via SharePoint REST to check if Graph incorrectly expanded EEEU as direct access..." -ForegroundColor Yellow
                }
                
                # Validate actual effective permissions using SharePoint REST API roleassignments
                # Graph API may report "write" but user only has Limited Access, or may expand EEEU as direct access
                Write-DebugInfo "Validating effective permissions via SharePoint REST API roleassignments..." -ForegroundColor Yellow
                $actualEEEURoles = @()
                try {
                    # Get the list item ID - need this to query roleassignments
                    $listItemId = $null
                    $libraryTitle = $null
                    
                    # Try to extract library name from drive name
                    if ($document.parentReference.driveId) {
                        $driveId = $document.parentReference.driveId
                        $drivesUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives"
                        $drives = Get-GraphAllPages -Uri $drivesUri -AccessToken $script:accessToken
                        $targetDrive = $drives | Where-Object { $_.id -eq $driveId }
                        if ($targetDrive) {
                            $libraryTitle = $targetDrive.name
                            Write-DebugInfo "  Library title: $libraryTitle" -ForegroundColor DarkGray
                        }
                    }
                    
                    # Get list item ID using the document's server-relative path
                    if ($document.webUrl) {
                        # Check if we already have the server-relative path and list item ID from initial retrieval
                        if ($document._spServerRelativePath -and $document.sharepointIds.listItemId) {
                            $fullServerRelativeUrl = $document._spServerRelativePath
                            $listItemId = $document.sharepointIds.listItemId
                            Write-DebugInfo "  Using cached server-relative path: $fullServerRelativeUrl" -ForegroundColor DarkGray
                            Write-DebugInfo "  Using cached list item ID: $listItemId" -ForegroundColor DarkGray
                        }
                        else {
                            # Build server-relative path from site path + library + item path
                            # Extract site path from site.webUrl
                            if ($site.webUrl -match 'https://[^/]+(/sites/[^/]+)') {
                                $sitePath = $matches[1]
                                
                                # Build the full server-relative path
                                # Library name mapping: "Documents" = "Shared Documents"
                                $libraryUrlSegment = if ($libraryTitle -eq "Documents") { "Shared Documents" } else { $libraryTitle }
                                
                                # Use parentReference.path if available, otherwise construct from document name
                                if ($document.parentReference.path) {
                                    # Path is like "/drives/{id}/root:/folder"
                                    # Extract the part after "root:"
                                    $folderPath = if ($document.parentReference.path -match 'root:(.*)$') { $matches[1] } else { "" }
                                    $fullServerRelativeUrl = "$sitePath/$libraryUrlSegment$folderPath/$($document.name)"
                                }
                                else {
                                    # Fallback: just use the document name
                                    $fullServerRelativeUrl = "$sitePath/$libraryUrlSegment/$($document.name)"
                                }
                                
                                Write-DebugInfo "  Server-relative path: $fullServerRelativeUrl" -ForegroundColor DarkGray
                                
                                # Use GetFileByServerRelativePath to get the list item
                                try {
                                    $endpoint = "web/GetFileByServerRelativePath(decodedUrl='$fullServerRelativeUrl')/ListItemAllFields"
                                    Write-DebugInfo "  Calling endpoint: $endpoint" -ForegroundColor DarkGray
                                    $listItem = Invoke-SharePointRestRequest -SiteUrl $site.webUrl -Endpoint $endpoint -AccessToken $script:sharePointToken
                                    if ($listItem.d) {
                                        $listItemId = $listItem.d.ID
                                        Write-DebugInfo "  ✓ Found List Item ID: $listItemId" -ForegroundColor Green
                                    }
                                    else {
                                        Write-DebugInfo "  ⚠ List item response has no .d property" -ForegroundColor DarkYellow
                                    }
                                }
                                catch {
                                    Write-DebugInfo "  ✗ Could not get list item via server-relative path: $($_.Exception.Message)" -ForegroundColor Red
                                }
                            }
                        }
                    }
                    
                    # Now check roleassignments if we have the required info
                    if ($listItemId -and $libraryTitle) {
                        Write-DebugInfo "  Checking roleassignments for item $listItemId in library '$libraryTitle'" -ForegroundColor DarkGray
                        $endpoint = "web/lists/GetByTitle('$libraryTitle')/items($listItemId)/roleassignments?`$expand=Member,RoleDefinitionBindings"
                        
                        $roleAssignments = Invoke-SharePointRestRequest -SiteUrl $site.webUrl -Endpoint $endpoint -AccessToken $script:sharePointToken
                        if ($roleAssignments.d.results) {
                            Write-DebugInfo "  Found $($roleAssignments.d.results.Count) role assignment(s)" -ForegroundColor DarkGray
                            
                            # Check if user has a direct role assignment
                            $userLoginName = "i:0#.f|membership|$($userToCheck.ToLower())"
                            $userAssignment = $roleAssignments.d.results | Where-Object { 
                                $_.Member.LoginName -eq $userLoginName 
                            }
                            
                            if ($userAssignment) {
                                $userRoles = $userAssignment.RoleDefinitionBindings.results | ForEach-Object { $_.Name }
                                Write-DebugInfo "  ✓ Found direct role assignment: $($userRoles -join ', ')" -ForegroundColor Green
                                
                                # Check if user ONLY has Limited Access
                                $nonLimitedRoles = $userRoles | Where-Object { $_ -ne "Limited Access" }
                                if ($nonLimitedRoles.Count -eq 0) {
                                    Write-DebugInfo "  ⚠ User only has Limited Access role" -ForegroundColor Yellow
                                    # Don't set actualEEEURoles yet - still need to check EEEU
                                }
                                else {
                                    # User has meaningful direct access - track it
                                    if ($userRoles -contains "Full Control") {
                                        $accessType += "; Direct Document Access (owner)"
                                        $hasMeaningfulDirectAccess = $true
                                        Write-DebugInfo "  ✓ Confirmed: User has Full Control permission" -ForegroundColor Green
                                    }
                                    elseif ($userRoles -contains "Edit" -or $userRoles -contains "Contribute") {
                                        $accessType += "; Direct Document Access (write)"
                                        $hasMeaningfulDirectAccess = $true
                                        Write-DebugInfo "  ✓ Confirmed: User has Edit/Contribute permission" -ForegroundColor Green
                                    }
                                    elseif ($userRoles -contains "Read") {
                                        $accessType += "; Direct Document Access (read)"
                                        $hasMeaningfulDirectAccess = $true
                                        Write-DebugInfo "  ✓ Confirmed: User has Read permission" -ForegroundColor Green
                                    }
                                }
                            }
                            else {
                                Write-DebugInfo "  No direct role assignment found for user" -ForegroundColor DarkGray
                            }
                            
                            # ALWAYS check if EEEU group has a role assignment (regardless of direct access)
                            # because user can have BOTH direct access AND EEEU access
                            # ALWAYS check if EEEU group has a role assignment (regardless of direct access)
                            # because user can have BOTH direct access AND EEEU access
                            $eeeuAssignment = $roleAssignments.d.results | Where-Object { 
                                $_.Member.LoginName -like "*spo-grid-all-users*"
                            }
                            
                            if ($eeeuAssignment) {
                                $eeeuRolesFromREST = $eeeuAssignment.RoleDefinitionBindings.results | ForEach-Object { $_.Name }
                                Write-DebugInfo "  ✓ Found EEEU group role assignment: $($eeeuRolesFromREST -join ', ')" -ForegroundColor Cyan
                                
                                # Check if EEEU ONLY has Limited Access
                                $nonLimitedRoles = $eeeuRolesFromREST | Where-Object { $_ -ne "Limited Access" }
                                if ($nonLimitedRoles.Count -eq 0) {
                                    Write-DebugInfo "  ⚠ EEEU group only has Limited Access role" -ForegroundColor Yellow
                                    $actualEEEURoles = @("Limited Access only")
                                }
                                else {
                                    # Map SharePoint role names to our internal format
                                    if ($eeeuRolesFromREST -contains "Full Control") {
                                        $actualEEEURoles = @("owner")
                                        Write-DebugInfo "  ✓ Confirmed: EEEU group grants Full Control permission" -ForegroundColor Green
                                    }
                                    elseif ($eeeuRolesFromREST -contains "Edit" -or $eeeuRolesFromREST -contains "Contribute") {
                                        $actualEEEURoles = @("write")
                                        Write-DebugInfo "  ✓ Confirmed: EEEU group grants Edit/Contribute permission" -ForegroundColor Green
                                    }
                                    elseif ($eeeuRolesFromREST -contains "Read") {
                                        $actualEEEURoles = @("read")
                                        Write-DebugInfo "  ✓ Confirmed: EEEU group grants Read permission" -ForegroundColor Green
                                    }
                                    else {
                                        # Has some other role that's not Limited Access
                                        $actualEEEURoles = @("read") # Default to read for safety
                                        Write-DebugInfo "  ✓ Confirmed: EEEU group grants permission: $($eeeuRolesFromREST -join ', ')" -ForegroundColor Green
                                    }
                                    # Set eeeuRoles if it wasn't set by Graph API
                                    if (-not $eeeuFound) {
                                        $eeeuRoles = $actualEEEURoles
                                        $eeeuFound = $true
                                        $userQualifiesForEEEU = $true
                                    }
                                    else {
                                        # Update with validated roles
                                        $eeeuRoles = $actualEEEURoles
                                    }
                                }
                            }
                            else {
                                Write-DebugInfo "  No EEEU group role assignment found" -ForegroundColor DarkGray
                            }
                            
                            # Handle the case where Graph API reported user access but SharePoint shows no meaningful direct assignment
                            # This means Graph incorrectly expanded EEEU permissions - clear the incorrect "Direct Access" flag
                            if ($hasMeaningfulDirectAccess -eq $false -and $userAssignment) {
                                # User has an assignment but it's only Limited Access
                                # Check if Graph API incorrectly reported this as direct access
                                $userRoles = $userAssignment.RoleDefinitionBindings.results | ForEach-Object { $_.Name }
                                $nonLimitedRoles = $userRoles | Where-Object { $_ -ne "Limited Access" }
                                if ($nonLimitedRoles.Count -eq 0) {
                                    # User only has Limited Access but Graph might have reported direct access
                                    # This is likely Graph incorrectly expanding EEEU
                                    Write-DebugInfo "  ⚠ User has Limited Access assignment - Graph may have incorrectly expanded EEEU as direct access" -ForegroundColor Yellow
                                }
                            }
                        }
                        else {
                            Write-DebugInfo "  No role assignments returned" -ForegroundColor DarkYellow
                        }
                    }
                    else {
                        Write-DebugInfo "  ⚠ Could not determine list item ID or library title for validation" -ForegroundColor DarkYellow
                        # Mark as validation failed since we couldn't even attempt the check
                        $actualEEEURoles = @("Validation Failed")
                    }
                }
                catch {
                    Write-DebugInfo "  Could not validate via roleassignments: $($_.Exception.Message)" -ForegroundColor DarkYellow
                    # If we can't validate, we'll use Graph API roles but mark as unvalidated
                    Write-DebugInfo "  ⚠ Cannot validate effective permissions - using Graph API reported roles (unvalidated)" -ForegroundColor Yellow
                    $actualEEEURoles = @() # Empty means use Graph API roles
                }
                
                # Use actual effective roles if validation succeeded, otherwise use Graph API roles
                if ($actualEEEURoles.Count -gt 0) {
                    if ($actualEEEURoles -notcontains "Validation Failed") {
                        # Validation succeeded - use the validated roles
                        $eeeuRoles = $actualEEEURoles
                    }
                    # If "Validation Failed", keep the original Graph API roles
                }
                
                # Report EEEU if user qualifies and EEEU has meaningful access
                # Report even if user has direct access - both can coexist
                if ($userQualifiesForEEEU -and $eeeuFound) {
                    # Don't report if EEEU only has Limited Access
                    if ($eeeuRoles -contains "Limited Access only") {
                        Write-DebugInfo "✗ EEEU permission exists but only has Limited Access - NOT REPORTING EEEU" -ForegroundColor Red
                    }
                    else {
                        # Check if we already added EEEU to accessType (to avoid duplicates)
                        if ($accessType -notlike "*Via Everyone Except External Users*") {
                            $userFound = $true
                            # Add EEEU access to access type
                            $eeeuAccessType = "Via Everyone Except External Users ($(($eeeuRoles | Select-Object -Unique | Sort-Object) -join ', '))"
                            if ($actualEEEURoles -contains "Validation Failed") {
                                $eeeuAccessType += " [Unvalidated - REST API unavailable]"
                            }
                            $accessType += "; $eeeuAccessType"
                            Write-DebugInfo "✓ User has access via 'Everyone except external users' with effective permissions: $(($eeeuRoles -join ', '))" -ForegroundColor Green
                        }
                        else {
                            Write-DebugInfo "  Skipping duplicate EEEU entry" -ForegroundColor DarkGray
                        }
                    }
                }
            }
            else {
                Write-DebugInfo "User is external and does NOT qualify for 'Everyone except external users' access" -ForegroundColor DarkYellow
            }
        }
        #endregion Effective Permission Validation
        
        #region CSV Output Generation
        if ($userFound) {
            #Collecting Export Properties for CSV File
            $ExportItem = New-Object PSObject
            $ExportItem  | Add-Member -MemberType NoteProperty -name "SiteName" -value $site.displayName
            $ExportItem  | Add-Member -MemberType NoteProperty -name "URL" -value $site.webUrl
            $ExportItem  | Add-Member -MemberType NoteProperty -name "DocumentName" -value $document.name
            $ExportItem  | Add-Member -MemberType NoteProperty -name "DocumentURL" -value $document.webUrl
            $ExportItem  | Add-Member -MemberType NoteProperty -name "User" -value $userToCheck
            
            # Get document owner from createdBy
            $ownerDisplay = "N/A"
            if ($document.createdBy -and $document.createdBy.user) {
                $displayName = $document.createdBy.user.displayName
                $email = $document.createdBy.user.email
                
                if ($displayName -and $email) {
                    # Both name and email available
                    $ownerDisplay = "$displayName ($email)"
                }
                elseif ($displayName) {
                    # Only name available
                    $ownerDisplay = "$displayName ()"
                }
                elseif ($email) {
                    # Only email available
                    $ownerDisplay = $email
                }
            }
            $ExportItem  | Add-Member -MemberType NoteProperty -name "Owner" -value $ownerDisplay
            
            # Clean up AccessType by removing any entries that are ONLY "Limited Access" (not descriptive text containing that phrase)
            $cleanAccessType = ($accessType.TrimStart('; ').Trim() -split '; ') | Where-Object { $_ -ne "Limited Access" }
            $finalAccessType = $cleanAccessType -join '; '
            
            $ExportItem  | Add-Member -MemberType NoteProperty -name "AccessType" -value $finalAccessType
            $siteOutput += $ExportItem
            
            Write-StatusMessage "✓ SUCCESS: Found $userToCheck on document '$($document.name)' with access: $finalAccessType" -ForegroundColor Green
            Write-LogEntry -LogName:$Log -LogEntryText "Found $userToCheck on document '$($document.name)' in site '$($site.webUrl)' with access: $finalAccessType"
        } 
        else {
            # Add "No Access" entry to CSV
            $ExportItem = New-Object PSObject
            $ExportItem | Add-Member -MemberType NoteProperty -name "SiteUrl" -value $site.webUrl
            $ExportItem | Add-Member -MemberType NoteProperty -name "SiteName" -value $site.displayName
            $ExportItem | Add-Member -MemberType NoteProperty -name "DocumentName" -value $document.name
            $ExportItem | Add-Member -MemberType NoteProperty -name "DocumentUrl" -value $document.webUrl
            $ExportItem | Add-Member -MemberType NoteProperty -name "User" -value $userToCheck
            $ExportItem | Add-Member -MemberType NoteProperty -name "AccessType" -value "No Access Found"
            $siteOutput += $ExportItem
            
            Write-StatusMessage "✗ NOT FOUND: $userToCheck does not have access to document '$($document.name)'" -ForegroundColor Red
            Write-LogEntry -LogName:$Log -LogEntryText "$userToCheck WAS NOT FOUND on document '$($document.name)' in site '$($site.webUrl)'"
        }
    }
}
catch {
    Write-StatusMessage "Error checking user: $($_.Exception.Message)" -ForegroundColor Red
    Write-LogEntry -LogName:$Log -LogEntryText "Error checking $userToCheck on '$($site.webUrl)' - Error: $($_.Exception.Message)"
}
#endregion CSV Output Generation
#endregion Permission Checking

#region Final Output
# Write output to CSV file (always write, even if no access found)
if ($siteOutput.Count -gt 0) {
    $siteOutput | Export-Csv $outputfile -NoTypeInformation
    Write-StatusMessage "Exported results to CSV" -ForegroundColor Green
    Write-LogEntry -LogName:$Log -LogEntryText "Exported results for site '$($site.displayName)' to CSV"
}
else {
    Write-StatusMessage "No documents processed - no CSV output generated" -ForegroundColor Yellow
    Write-LogEntry -LogName:$Log -LogEntryText "No documents processed for site '$($site.displayName)'"
}

Write-StatusMessage "Completed processing site: $($site.displayName)" -ForegroundColor Yellow
Write-LogEntry -LogName:$Log -LogEntryText "Completed processing site: $($site.displayName)"

#Output Results
Write-Host ""
Write-StatusMessage "Processing completed!" -ForegroundColor Green
Write-StatusMessage "Output file saved to $outputfile" -ForegroundColor Green
Write-Host ""
Write-StatusMessage "Log file saved to $log" -ForegroundColor Green

# No disconnect needed for REST API - token expires automatically
#endregion Final Output
