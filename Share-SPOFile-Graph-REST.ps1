
#region Script Header
<#
.SYNOPSIS
Grants permissions to a user on a SharePoint/OneDrive file using Microsoft Graph API.

.DESCRIPTION
This script authenticates with Microsoft Graph API and uses the driveItem invite endpoint
to programmatically grant permissions (read, write, etc.) to a specified user on a file.
The file is specified via a direct SharePoint/OneDrive sharing link.

When the Graph API invite call fails (e.g., on Information Barrier enabled sites),
the script falls back to the SharePoint REST API to grant permissions directly via
SharePoint's role assignment endpoint

.PARAMETER None
This script does not accept parameters through the command line. Configuration is done through variables
at the beginning of the script.

.NOTES
File Name       : Share-SPOFile-Graph-REST.ps1
Author          : Mike Lee
Date Created    : 7/18/25
Date Updated    : 3/5/26 (with SharePoint REST API fallback)
Prerequisites   : 
- PowerShell 7.4 or higher
- Appropriate permissions in Azure AD 

API Permissions Required:
- Sites.ReadWrite.All (for SharePoint sites)
- Files.ReadWrite.All (for file access and sharing)

- Microsoft Graph API access

.EXAMPLE
PS> .\Share-SPOFile-Graph-REST.ps1
Grants the configured permissions to the specified user on the target file.

.OUTPUTS
Console output indicating success or failure of the permission grant.

.LINK
https://learn.microsoft.com/en-us/graph/api/driveitem-invite?view=graph-rest-1.0&tabs=http

.COMPONENT
Microsoft Graph API

.FUNCTIONALITY
- Authenticates with Microsoft Graph API using client credentials or certificate
- Parses SharePoint/OneDrive sharing links to extract site and item IDs
- Grants permissions to users via the driveItem invite API
- Falls back to SharePoint REST API for sites where Graph invite is blocked (e.g., IB-enabled sites)
- Handles throttling using exponential backoff
#>
#endregion Script Header

#region Configuration
##############################################################
#                  CONFIGURATION SECTION                    #
#############################################################
# Modify these values according to your environment

#-----------------------------------------------------------#
#              FILE SHARING CONFIGURATION                   #
#-----------------------------------------------------------#

# Direct link to the SharePoint/OneDrive file to share
# This can be a sharing link or direct URL to the file
# Examples:
#   - https://contoso.sharepoint.com/:x:/s/SiteName/EaBcDeFgHiJkLmNoPqRsTuVw?e=AbCdEf
#   - https://contoso.sharepoint.com/sites/SiteName/_layouts/15/Doc.aspx?sourcedoc={guid}&file=filename.xlsx
$fileLink = "https://m365cpi13246019.sharepoint.com/sites/TeamSiteNoGroup1/Shared%20Documents/Testdoc1.docx"

# Alternatively, you can specify the site ID and item ID directly (leave $fileLink empty to use these)
$siteId = ""    # Example: "490b8ce0-7724-439c-8f63-eb7a881d784d"
$itemId = ""    # Example: "b!4IwLSSR3nEOPY-t6iB14TXArkmW_fN9Nq51uVS5gVgi2jMS4JnfYSqk6u2ieXO_H"

#-----------------------------------------------------------#
#              INVITATION/PERMISSION SETTINGS               #
#-----------------------------------------------------------#

# Email address of the user to grant access to
$email = "PeytonD@M365CPI13246019.onmicrosoft.com"

# Message to include with the invitation (optional)
$message = "Granting access for collaboration"

# Whether the recipient must sign in to access the file
$requireSignIn = $true

# Whether to send an email invitation to the recipient
$sendInvitation = $false

# Roles to grant - valid values: "read", "write", "owner"
# Can be a single role or multiple roles as an array
$roles = @("read")

#-----------------------------------------------------------#
#              AUTHENTICATION CONFIGURATION                 #
#-----------------------------------------------------------#

# Enable or disable verbose debug output
# Set to $true for detailed logging, $false for basic info only
$debug = $false

# Set the tenant ID, client ID, and client secret for authentication
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3';
$clientId = '1e488dc4-1977-48ef-8d4d-9856f4e04536';

# Authentication type: Choose 'ClientSecret' or 'Certificate'
$AuthType = 'Certificate'  # Valid values: 'ClientSecret' or 'Certificate'

# Client Secret authentication (used when $AuthType = 'ClientSecret')
$clientSecret = '';

# Certificate authentication (used when $AuthType = 'Certificate')
$Thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"

# Certificate store location: Choose 'LocalMachine' or 'CurrentUser'
$CertStore = 'LocalMachine'  # Valid values: 'LocalMachine' or 'CurrentUser'

# REST Fallback Configuration
# When Graph API fails to set permissions (e.g., on IB-enabled sites),
# the script will fall back to the SharePoint REST API to grant permissions directly.
# uses only native Invoke-RestMethod calls.
$useRESTFallback = $true  # Set to $false to disable REST fallback

#############################################################
#                  END CONFIGURATION SECTION                #
#############################################################
#endregion Configuration

#region Initialization
# Load required assemblies
Add-Type -AssemblyName System.Web


# This ensures each log file has a unique name
$date = Get-Date -Format "yyyyMMddHHmmss";

# The log file will store the search results including sensitivity and retention labels in CSV format
$LogName = Join-Path -Path $env:TEMP -ChildPath ("SPOFileswithLabels_Search_Results_" + $date + ".csv");

# Initialize global variables for the token and search results
$global:token = @();
$global:tokenExpiry = $null;
$global:spToken = $null;       # SharePoint-scoped token for REST fallback
$global:spTokenExpiry = $null; # SharePoint token expiry
$global:spTokenSite = $null;   # Tracks which site the SP token was issued for
$global:Results = @();
#endregion Initialization

#region Graph Throttle Handling
# Function to handle throttling for Microsoft Graph requests
# This implements best practices from https://learn.microsoft.com/en-us/graph/throttling
# It automatically handles 429 "Too Many Requests" responses with either:
# 1. The Retry-After header value if provided by the server
# 2. Exponential backoff if no Retry-After header is present
function Invoke-GraphRequestWithThrottleHandling {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [string]$Method,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers = @{},
        
        [Parameter(Mandatory = $false)]
        [string]$Body = $null,
        
        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json",
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 15,  # Increased from 10
        
        [Parameter(Mandatory = $false)]
        [int]$InitialBackoffSeconds = 3,  # Increased from 2
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 300  # 5 minute timeout
    )
    
    $retryCount = 0
    $backoffSeconds = $InitialBackoffSeconds
    $success = $false
    $result = $null
    
    if ($debug) {
        Write-Host "Making Graph request to $Uri" -ForegroundColor Gray
    }
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            # Create web request with timeout
            if ($Body) {
                $result = Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -Body $Body -ContentType $ContentType -TimeoutSec $TimeoutSeconds -ErrorAction Stop
            }
            else {
                $result = Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -ContentType $ContentType -TimeoutSec $TimeoutSeconds -ErrorAction Stop
            }
            $success = $true
        }
        catch [System.Net.WebException] {
            $webException = $_.Exception
            $statusCode = $null
            
            # Handle different types of web exceptions
            if ($webException.Response) {
                $statusCode = [int]$webException.Response.StatusCode
            }
            
            # Check for timeout or connection errors
            if ($webException.Status -eq [System.Net.WebExceptionStatus]::Timeout -or 
                $webException.Status -eq [System.Net.WebExceptionStatus]::ConnectionClosed -or
                $webException.Status -eq [System.Net.WebExceptionStatus]::ConnectFailure -or
                $statusCode -eq 502 -or $statusCode -eq 503 -or $statusCode -eq 504) {
                
                $retryCount++
                $waitTime = [Math]::Min($backoffSeconds, 300)  # Cap at 5 minutes
                
                Write-Host "Connection/timeout error detected. Status: $($webException.Status). Waiting $waitTime seconds before retry. Attempt $retryCount of $MaxRetries..." -ForegroundColor Yellow
                
                if ($retryCount -lt $MaxRetries) {
                    Start-Sleep -Seconds $waitTime
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 300)  # Exponential backoff capped at 5 minutes
                }
                else {
                    Write-Host "Maximum retry attempts reached ($MaxRetries). Giving up." -ForegroundColor Red
                    throw $_
                }
            }
            elseif ($statusCode -eq 429) {
                # Handle throttling
                $retryAfter = $null
                if ($webException.Response.Headers["Retry-After"]) {
                    $retryAfter = [int]($webException.Response.Headers["Retry-After"])
                    Write-Host "Request throttled. Retry-After header suggests waiting for $retryAfter seconds." -ForegroundColor Yellow
                }
                else {
                    $retryAfter = $backoffSeconds
                    Write-Host "Request throttled. Using exponential backoff: waiting for $retryAfter seconds." -ForegroundColor Yellow
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 300)
                }
                
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Host "Throttling detected. Waiting before retry. Attempt $retryCount of $MaxRetries..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfter
                }
                else {
                    Write-Host "Maximum retry attempts reached ($MaxRetries). Giving up." -ForegroundColor Red
                    throw $_
                }
            }
            else {
                # Not a recoverable error, rethrow
                throw $_
            }
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = $_.Exception.Response.StatusCode.value__
            }
            
            # Check if this is a throttling error (429) or server error (5xx)
            if ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -le 599)) {
                # Get the Retry-After header if it exists
                $retryAfter = $null
                if ($statusCode -eq 429 -and $_.Exception.Response.Headers.Contains("Retry-After")) {
                    $retryAfter = [int]($_.Exception.Response.Headers.GetValues("Retry-After") | Select-Object -First 1)
                    Write-Host "Request throttled. Retry-After header suggests waiting for $retryAfter seconds." -ForegroundColor Yellow
                }
                else {
                    # If no Retry-After header, use exponential backoff
                    $retryAfter = $backoffSeconds
                    Write-Host "Server error ($statusCode) or throttling detected. Using exponential backoff: waiting for $retryAfter seconds." -ForegroundColor Yellow
                    # Increase backoff for next potential retry (exponential)
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 300)  # Cap at 5 minutes
                }
                
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Host "Retryable error detected. Waiting before retry. Attempt $retryCount of $MaxRetries..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfter
                }
                else {
                    Write-Host "Maximum retry attempts reached ($MaxRetries). Giving up." -ForegroundColor Red
                    throw $_
                }
            }
            else {
                # Not a throttling error, rethrow
                throw $_
            }
        }
    }
    
    return $result
}
#endregion Graph Throttle Handling

#region Authentication
# This function authenticates with Microsoft Graph API and retrieves an access token
function AcquireToken() {
    Write-Host "Connecting to Microsoft Graph using $AuthType authentication..." -ForegroundColor Cyan
    
    if ($AuthType -eq 'ClientSecret') {
        # Client Secret authentication
        $uri = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/token";
        
        # Define the body for the authentication request
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $clientId
            client_secret = $clientSecret
            resource      = 'https://graph.microsoft.com'
            scope         = 'https://graph.microsoft.com/.default'
        };
        
        try {
            # Send the authentication request and extract the token
            $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop;
            $global:token = $loginResponse.access_token;
            
            # Calculate token expiry (typically 1 hour, but we'll refresh before then)
            $expiresIn = if ($loginResponse.expires_in) { $loginResponse.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)  # Refresh 5 minutes before expiry
            
            Write-Host "Successfully connected using Client Secret authentication. Token expires at: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to connect using Client Secret authentication" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }
    }
    elseif ($AuthType -eq 'Certificate') {
        # Certificate authentication
        $uri = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/v2.0/token";
        
        # Get the certificate from the local certificate store
        try {
            $cert = Get-Item -Path "Cert:\$CertStore\My\$Thumbprint" -ErrorAction Stop
        }
        catch {
            Write-Host "Certificate with thumbprint $Thumbprint not found in $CertStore\My store" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }
        
        # Create the JWT assertion for certificate authentication
        $now = [System.DateTimeOffset]::UtcNow
        $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        $nbf = $now.ToUnixTimeSeconds()
        $aud = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
        
        # Create JWT header
        $header = @{
            alg = "RS256"
            typ = "JWT"
            x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        } | ConvertTo-Json -Compress
        
        # Create JWT payload
        $payload = @{
            aud = $aud
            exp = $exp
            iss = $clientId
            jti = [System.Guid]::NewGuid().ToString()
            nbf = $nbf
            sub = $clientId
        } | ConvertTo-Json -Compress
        
        # Base64 encode header and payload
        $headerBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $payloadBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        # Create the string to sign
        $stringToSign = "$headerBase64.$payloadBase64"
        
        # Sign the string with the certificate
        $signature = $cert.PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($stringToSign), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $signatureBase64 = [Convert]::ToBase64String($signature).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        # Create the final JWT
        $jwt = "$stringToSign.$signatureBase64"
        
        # Define the body for the authentication request
        $body = @{
            client_id             = $clientId
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = $jwt
            scope                 = "https://graph.microsoft.com/.default"
            grant_type            = "client_credentials"
        }
        
        try {
            # Send the authentication request and extract the token
            $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop;
            $global:token = $loginResponse.access_token;
            
            # Calculate token expiry (typically 1 hour, but we'll refresh before then)
            $expiresIn = if ($loginResponse.expires_in) { $loginResponse.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)  # Refresh 5 minutes before expiry
            
            Write-Host "Successfully connected using Certificate authentication. Token expires at: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to connect using Certificate authentication" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }
    }
    else {
        Write-Host "Invalid authentication type: $AuthType. Valid values are 'ClientSecret' or 'Certificate'." -ForegroundColor Red
        Exit
    }
}
#endregion Authentication

#region Token Validation
# Function to check if token needs refresh and refresh if necessary
function Test-ValidToken() {
    if ($null -eq $global:tokenExpiry -or (Get-Date) -gt $global:tokenExpiry) {
        Write-Host "Token expired or expiring soon. Refreshing..." -ForegroundColor Yellow
        AcquireToken
    }
}
#endregion Token Validation

#region Get-DriveItemFromLink
# Function to parse a SharePoint/OneDrive sharing link and extract site and item information
function Get-DriveItemFromLink {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SharingLink
    )
    
    Write-Host "Parsing sharing link to extract drive item information..." -ForegroundColor Cyan
    
    if ($debug) {
        Write-Host "  Input link: $SharingLink" -ForegroundColor Gray
    }
    
    try {
        # Ensure we have a valid token
        Test-ValidToken
        
        $headers = @{
            "Authorization" = "Bearer $global:token"
            "Content-Type"  = "application/json"
        }
        
        # Encode the sharing URL using the sharing URL encoding format
        # https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0
        $base64Value = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($SharingLink))
        $encodedUrl = "u!" + $base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-')
        
        if ($debug) {
            Write-Host "  Encoded URL: $encodedUrl" -ForegroundColor Gray
        }
        
        # Use the shares endpoint to decode the sharing link and get the driveItem
        $sharesUri = "https://graph.microsoft.com/v1.0/shares/$encodedUrl/driveItem"
        
        if ($debug) {
            Write-Host "  Calling shares API: $sharesUri" -ForegroundColor Gray
        }
        
        $driveItem = Invoke-GraphRequestWithThrottleHandling -Uri $sharesUri -Method "GET" -Headers $headers
        
        if ($driveItem) {
            # Extract the parentReference which contains siteId and driveId
            $siteId = $driveItem.parentReference.siteId
            $driveId = $driveItem.parentReference.driveId
            $itemId = $driveItem.id
            $fileName = $driveItem.name
            
            # Build the actual server-relative file path by querying the drive's webUrl
            # parentReference.path is library-relative (e.g. /drives/{id}/root:/subfolder)
            # but does NOT include the site or library path, so we must get the drive's webUrl
            # which gives us the full library server-relative path
            $serverRelativeUrl = $null
            $resolvedSiteUrl = $null

            try {
                # Get the drive info to find the library's server-relative path
                # drive.webUrl = e.g. "https://tenant.sharepoint.com/sites/sitename/Shared Documents"
                $driveInfoUri = "https://graph.microsoft.com/v1.0/drives/$driveId"
                $driveInfo = Invoke-GraphRequestWithThrottleHandling -Uri $driveInfoUri -Method "GET" -Headers $headers

                if ($driveInfo.webUrl) {
                    $driveWebUri = [System.Uri]$driveInfo.webUrl
                    $libraryServerRelativePath = [System.Web.HttpUtility]::UrlDecode($driveWebUri.AbsolutePath)

                    # Get the folder path within the library from parentReference.path (after "root:")
                    $folderInLib = ""
                    if ($driveItem.parentReference.path -match "root:(.+)$") {
                        $folderInLib = $matches[1]
                    }

                    # Build the full server-relative URL: /sites/sitename/Shared Documents[/subfolder]/filename
                    $serverRelativeUrl = $libraryServerRelativePath.TrimEnd('/') + $folderInLib.TrimEnd('/') + "/" + $fileName
                    Write-Host "  Server-Relative URL: $serverRelativeUrl" -ForegroundColor Cyan

                    # Extract site URL from the drive's webUrl
                    $hostUrl = $driveWebUri.Scheme + "://" + $driveWebUri.Host
                    if ($libraryServerRelativePath -match "^(/sites/[^/]+|/teams/[^/]+)") {
                        $resolvedSiteUrl = $hostUrl + $matches[1]
                    }
                    else {
                        $resolvedSiteUrl = $hostUrl
                    }
                    Write-Host "  Site URL: $resolvedSiteUrl" -ForegroundColor Cyan
                }
            }
            catch {
                Write-Host "  Warning: Could not resolve drive webUrl from Graph: $($_.Exception.Message)" -ForegroundColor Yellow
            }

            # Fallback: extract site URL from driveItem.webUrl if not already resolved
            if (-not $resolvedSiteUrl -and $driveItem.webUrl) {
                $webUri = [System.Uri]$driveItem.webUrl
                $hostUrl = $webUri.Scheme + "://" + $webUri.Host
                $webPath = $webUri.AbsolutePath
                if ($webPath -match "^(/sites/[^/]+|/teams/[^/]+)") {
                    $resolvedSiteUrl = $hostUrl + $matches[1]
                }
                else {
                    $resolvedSiteUrl = $hostUrl
                }
                Write-Host "  Site URL (from webUrl fallback): $resolvedSiteUrl" -ForegroundColor Cyan
            }

            Write-Host "✓ Successfully resolved sharing link" -ForegroundColor Green
            Write-Host "  File Name: $fileName" -ForegroundColor Cyan
            Write-Host "  Site ID: $siteId" -ForegroundColor Cyan
            Write-Host "  Drive ID: $driveId" -ForegroundColor Cyan
            Write-Host "  Item ID: $itemId" -ForegroundColor Cyan
            
            return @{
                SiteId            = $siteId
                DriveId           = $driveId
                ItemId            = $itemId
                FileName          = $fileName
                DriveItem         = $driveItem
                ServerRelativeUrl = $serverRelativeUrl
                SiteUrl           = $resolvedSiteUrl
            }
        }
        else {
            Write-Host "❌ Failed to resolve sharing link - no drive item returned" -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "❌ Failed to parse sharing link" -ForegroundColor Red
        Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Yellow
        
        if ($debug) {
            Write-Host "   Full error: $($_.Exception)" -ForegroundColor Gray
        }
        
        return $null
    }
}
#endregion Get-DriveItemFromLink

#region Grant-DriveItemPermission (Graph API)
# Function to grant permissions to a user on a drive item using the invite API
function Grant-DriveItemPermission {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteId,
        
        [Parameter(Mandatory = $true)]
        [string]$DriveId,
        
        [Parameter(Mandatory = $true)]
        [string]$ItemId,
        
        [Parameter(Mandatory = $true)]
        [string]$Email,
        
        [Parameter(Mandatory = $false)]
        [string]$Message = "Granting access",
        
        [Parameter(Mandatory = $false)]
        [bool]$RequireSignIn = $true,
        
        [Parameter(Mandatory = $false)]
        [bool]$SendInvitation = $false,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Roles = @("read")
    )
    
    Write-Host "`nGranting permissions to user..." -ForegroundColor Cyan
    Write-Host "  Recipient: $Email" -ForegroundColor White
    Write-Host "  Roles: $($Roles -join ', ')" -ForegroundColor White
    Write-Host "  Require Sign-In: $RequireSignIn" -ForegroundColor White
    Write-Host "  Send Invitation: $SendInvitation" -ForegroundColor White
    
    try {
        # Ensure we have a valid token
        Test-ValidToken
        
        $headers = @{
            "Authorization" = "Bearer $global:token"
            "Content-Type"  = "application/json"
        }
        
        # Build the invite request body
        $inviteBody = @{
            recipients     = @(
                @{
                    email = $Email
                }
            )
            message        = $Message
            requireSignIn  = $RequireSignIn
            sendInvitation = $SendInvitation
            roles          = $Roles
        } | ConvertTo-Json -Depth 10
        
        if ($debug) {
            Write-Host "  Request body: $inviteBody" -ForegroundColor Gray
        }
        
        # Construct the invite API URI using /drives/{driveId}/items/{itemId}/invite
        # Reference: https://learn.microsoft.com/en-us/graph/api/driveitem-invite?view=graph-rest-1.0
        $inviteUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/invite"
        
        Write-Host "  API URI: $inviteUri" -ForegroundColor Gray
        if ($debug) {
            Write-Host "  Request body: $inviteBody" -ForegroundColor Gray
        }
        
        $response = Invoke-RestMethod -Uri $inviteUri -Method Post -Headers $headers -Body $inviteBody -ContentType "application/json" -ErrorAction Stop
        
        if ($response -and $response.value) {
            Write-Host "`n✓ Permission granted successfully!" -ForegroundColor Green
            
            foreach ($permission in $response.value) {
                Write-Host "  Permission ID: $($permission.id)" -ForegroundColor Cyan
                Write-Host "  Roles: $($permission.roles -join ', ')" -ForegroundColor Cyan
                
                if ($permission.grantedTo) {
                    Write-Host "  Granted To: $($permission.grantedTo.user.displayName) ($($permission.grantedTo.user.email))" -ForegroundColor Cyan
                }
                elseif ($permission.grantedToIdentities) {
                    foreach ($identity in $permission.grantedToIdentities) {
                        Write-Host "  Granted To: $($identity.user.displayName) ($($identity.user.email))" -ForegroundColor Cyan
                    }
                }
            }
            
            return $response
        }
        else {
            Write-Host "✓ Permission request completed (no detailed response returned)" -ForegroundColor Green
            return $response
        }
    }
    catch {
        $statusCode = $null
        $errorMessage = $_.Exception.Message
        $responseBody = $null
        
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }
        
        # Try to read the response body for more details
        # PowerShell stores the error response in $_.ErrorDetails
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $responseBody = $_.ErrorDetails.Message
            try {
                $errorDetails = $responseBody | ConvertFrom-Json
                if ($errorDetails.error.message) {
                    $errorMessage = $errorDetails.error.message
                }
                if ($errorDetails.error.code) {
                    $errorMessage = "[$($errorDetails.error.code)] $errorMessage"
                }
            }
            catch {
                # If JSON parsing fails, use the raw response body
                $errorMessage = $responseBody
            }
        }
        
        Write-Host "`n❌ Failed to grant permission" -ForegroundColor Red
        Write-Host "   Status Code: $statusCode" -ForegroundColor Yellow
        Write-Host "   Error: $errorMessage" -ForegroundColor Yellow
        
        if ($responseBody) {
            Write-Host "   Response Body: $responseBody" -ForegroundColor Gray
        }
        
        if ($debug) {
            Write-Host "   Full error: $($_.Exception)" -ForegroundColor Gray
        }
        
        return $null
    }
}
#endregion Grant-DriveItemPermission (Graph API)

#region REST Fallback Functions
#############################################################
#        REST FALLBACK FUNCTIONS (IB-ENABLED SITES)         #
#  Uses SharePoint REST API only                            #
#############################################################

#region REST Throttle Handling
# Function to handle throttling for SharePoint REST API operations
function Invoke-SPRestWithThrottleHandling {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [string]$Method = "GET",

        [Parameter(Mandatory = $false)]
        [hashtable]$Headers = @{},

        [Parameter(Mandatory = $false)]
        [string]$Body = $null,

        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json;odata=nometadata",

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 10,

        [Parameter(Mandatory = $false)]
        [int]$InitialBackoffSeconds = 3
    )

    $retryCount = 0
    $backoffSeconds = $InitialBackoffSeconds

    while ($retryCount -lt $MaxRetries) {
        try {
            $params = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $Headers
                ErrorAction = "Stop"
            }
            if ($Body) {
                $params['Body'] = $Body
                $params['ContentType'] = $ContentType
            }

            if ($debug) {
                Write-Host "    SP REST: $Method $Uri" -ForegroundColor Gray
            }

            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            $retryAfter = $backoffSeconds
            if ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -le 599)) {
                # Honor Retry-After header when present
                if ($_.Exception.Response -and $_.Exception.Response.Headers["Retry-After"]) {
                    $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                }
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Host "  SP REST throttle/error ($statusCode). Retrying in $retryAfter seconds... (Attempt $retryCount of $MaxRetries)" -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfter
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 120)
                }
                else {
                    Write-Host "  SP REST maximum retries reached ($MaxRetries). Giving up." -ForegroundColor Red
                    throw $_
                }
            }
            else {
                throw $_
            }
        }
    }
}
#endregion REST Throttle Handling

#region REST Helper Functions
# Function to extract the SharePoint site URL from a file URL
function Get-SiteUrlFromFileUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FileUrl
    )

    try {
        $uri = [System.Uri]$FileUrl
        $hostUrl = $uri.Scheme + "://" + $uri.Host
        $path = $uri.AbsolutePath

        # Match /sites/xxx or /teams/xxx pattern
        if ($path -match "^(/sites/[^/]+|/teams/[^/]+)") {
            return $hostUrl + $matches[1]
        }
        else {
            # Root site
            return $hostUrl
        }
    }
    catch {
        Write-Host "  Could not parse site URL from: $FileUrl - $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Acquires a SharePoint-scoped OAuth token using the same credentials as the Graph token.
# SharePoint REST API requires a token scoped to the SharePoint host (not graph.microsoft.com).
function Get-SharePointToken {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    # Derive the SharePoint host resource URL (e.g. https://tenant.sharepoint.com)
    $spUri = [System.Uri]$SiteUrl
    $spBaseUrl = $spUri.Scheme + "://" + $spUri.Host

    # Return cached token if still valid and for the same site
    if ($global:spToken -and $global:spTokenExpiry -and (Get-Date) -lt $global:spTokenExpiry -and $global:spTokenSite -eq $spBaseUrl) {
        if ($debug) { Write-Host "  Using cached SharePoint token (expires $($global:spTokenExpiry))" -ForegroundColor Gray }
        return $global:spToken
    }

    Write-Host "  Acquiring SharePoint-scoped token for: $spBaseUrl" -ForegroundColor Cyan

    try {
        if ($AuthType -eq 'ClientSecret') {
            $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/token"
            $body = @{
                grant_type    = "client_credentials"
                client_id     = $clientId
                client_secret = $clientSecret
                resource      = $spBaseUrl
            }
            $response = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        }
        elseif ($AuthType -eq 'Certificate') {
            $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
            $cert = Get-Item -Path "Cert:\$CertStore\My\$Thumbprint" -ErrorAction Stop

            # Build the same JWT assertion used in AcquireToken, but targeting the SP scope
            $now = [System.DateTimeOffset]::UtcNow
            $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
            $nbf = $now.ToUnixTimeSeconds()
            $aud = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

            $header = @{
                alg = "RS256"
                typ = "JWT"
                x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+', '-').Replace('/', '_')
            } | ConvertTo-Json -Compress

            $payload = @{
                aud = $aud
                exp = $exp
                iss = $clientId
                jti = [System.Guid]::NewGuid().ToString()
                nbf = $nbf
                sub = $clientId
            } | ConvertTo-Json -Compress

            $headerB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
            $payloadB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
            $strToSign = "$headerB64.$payloadB64"
            $sig = $cert.PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($strToSign), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            $sigB64 = [Convert]::ToBase64String($sig).TrimEnd('=').Replace('+', '-').Replace('/', '_')
            $jwt = "$strToSign.$sigB64"

            $body = @{
                client_id             = $clientId
                client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
                client_assertion      = $jwt
                scope                 = "$spBaseUrl/.default"
                grant_type            = "client_credentials"
            }
            $response = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        }
        else {
            throw "Unsupported AuthType '$AuthType' for SharePoint token acquisition."
        }

        $expiresIn = if ($response.expires_in) { $response.expires_in } else { 3600 }
        $global:spToken = $response.access_token
        $global:spTokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
        $global:spTokenSite = $spBaseUrl

        Write-Host "  ✓ SharePoint token acquired. Expires at: $($global:spTokenExpiry)" -ForegroundColor Green
        return $global:spToken
    }
    catch {
        Write-Host "  ❌ Failed to acquire SharePoint token: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}
#endregion REST Helper Functions

#region Grant-DriveItemPermissionViaREST
# Grants file-level permissions using the SharePoint REST API as a fallback when
# the Graph invite endpoint is unavailable (e.g., on Information Barrier enabled sites).
# Uses only native Invoke-RestMethod calls.
function Grant-DriveItemPermissionViaREST {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FileUrl,

        [Parameter(Mandatory = $true)]
        [string]$Email,

        [Parameter(Mandatory = $true)]
        [string[]]$Roles,

        [Parameter(Mandatory = $false)]
        [string]$ServerRelativeUrl = $null,

        [Parameter(Mandatory = $false)]
        [string]$SiteUrl = $null
    )

    Write-Host "`n  Attempting SharePoint REST fallback to grant permissions (for IB-enabled sites)..." -ForegroundColor Yellow
    Write-Host "  File URL: $FileUrl" -ForegroundColor White
    Write-Host "  Recipient: $Email" -ForegroundColor White
    Write-Host "  Roles: $($Roles -join ', ')" -ForegroundColor White

    # Map Graph API roles to SharePoint permission level names
    $permissionLevel = switch ($Roles[0]) {
        "read" { "Read" }
        "write" { "Edit" }
        "owner" { "Full Control" }
        default { "Read" }
    }
    Write-Host "  SharePoint permission level: $permissionLevel" -ForegroundColor Cyan

    # Resolve site URL
    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        $SiteUrl = Get-SiteUrlFromFileUrl -FileUrl $FileUrl
    }
    if (-not $SiteUrl) {
        Write-Host "  ❌ Could not determine site URL for REST fallback" -ForegroundColor Red
        return $null
    }

    # Resolve server-relative path of the file
    $relativePath = $null
    if (-not [string]::IsNullOrWhiteSpace($ServerRelativeUrl)) {
        $relativePath = [System.Web.HttpUtility]::UrlDecode($ServerRelativeUrl)
        Write-Host "  Server-relative path (from Graph driveItem): $relativePath" -ForegroundColor Cyan
    }
    else {
        $uri = [System.Uri]$FileUrl
        $rawPath = [System.Web.HttpUtility]::UrlDecode($uri.AbsolutePath)
        if ($rawPath -notmatch "_layouts/") {
            $relativePath = $rawPath
            Write-Host "  Server-relative path (from URL): $relativePath" -ForegroundColor Cyan
        }
        else {
            Write-Host "  ⚠ File URL is a _layouts URL - cannot use as file path." -ForegroundColor Yellow
        }
    }

    if ([string]::IsNullOrWhiteSpace($relativePath)) {
        Write-Host "  ❌ Could not determine the server-relative file path. Cannot set file-level permissions." -ForegroundColor Red
        return $null
    }

    # Acquire SharePoint-scoped token
    $spToken = $null
    try {
        $spToken = Get-SharePointToken -SiteUrl $SiteUrl
    }
    catch {
        Write-Host "  ❌ Cannot proceed without a SharePoint token: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    # Headers for all SharePoint REST calls
    $spHeaders = @{
        "Authorization" = "Bearer $spToken"
        "Accept"        = "application/json;odata=nometadata"
        "Content-Type"  = "application/json;odata=nometadata"
    }

    # Single-quote escape the relative path for use inside REST URL parameters
    $escapedRelPath = $relativePath.Replace("'", "''")

    # -------------------------------------------------------------------
    # STEP 1 – Break role inheritance on the list item (preserve existing)
    # -------------------------------------------------------------------
    try {
        Write-Host "  Step 1: Breaking role inheritance on file (preserving existing permissions)..." -ForegroundColor Cyan
        $breakUri = "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$escapedRelPath')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)"
        Invoke-SPRestWithThrottleHandling -Uri $breakUri -Method "POST" -Headers $spHeaders | Out-Null
        Write-Host "    ✓ Role inheritance broken (or was already unique)" -ForegroundColor Green
    }
    catch {
        # A 400 response here typically means inheritance was already broken - treat as non-fatal
        $sc = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
        if ($sc -eq 400) {
            Write-Host "    Note: Role inheritance may already be broken (HTTP 400 - continuing)" -ForegroundColor Gray
        }
        else {
            Write-Host "  ❌ Failed to break role inheritance: $($_.Exception.Message)" -ForegroundColor Red
            return $null
        }
    }

    # -------------------------------------------------------------------
    # STEP 2 – Ensure the user account exists in the site and get their ID
    # -------------------------------------------------------------------
    $userId = $null
    try {
        Write-Host "  Step 2: Resolving user account in site: $Email" -ForegroundColor Cyan
        $loginName = "i:0#.f|membership|$Email"
        $ensureUserUri = "$SiteUrl/_api/web/ensureuser"
        $ensureUserBody = '{"logonName":"' + $loginName + '"}'
        $userResponse = Invoke-SPRestWithThrottleHandling -Uri $ensureUserUri -Method "POST" -Headers $spHeaders -Body $ensureUserBody
        $userId = $userResponse.Id
        Write-Host "    ✓ User resolved. Site user ID: $userId" -ForegroundColor Green
    }
    catch {
        Write-Host "  ❌ Failed to resolve user '$Email' in site: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    # -------------------------------------------------------------------
    # STEP 3 – Get the role definition ID for the requested permission level
    # -------------------------------------------------------------------
    $roleDefId = $null
    try {
        Write-Host "  Step 3: Retrieving role definition for '$permissionLevel'..." -ForegroundColor Cyan
        $roleDefUri = "$SiteUrl/_api/web/roledefinitions/getbyname('$permissionLevel')"
        $roleDefResponse = Invoke-SPRestWithThrottleHandling -Uri $roleDefUri -Method "GET" -Headers $spHeaders
        $roleDefId = $roleDefResponse.Id
        Write-Host "    ✓ Role definition ID: $roleDefId" -ForegroundColor Green
    }
    catch {
        Write-Host "  ❌ Failed to retrieve role definition '$permissionLevel': $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    # -------------------------------------------------------------------
    # STEP 4 – Add the role assignment to the file's list item
    # -------------------------------------------------------------------
    try {
        Write-Host "  Step 4: Assigning '$permissionLevel' to user (ID $userId) on file..." -ForegroundColor Cyan
        $assignUri = "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$escapedRelPath')/ListItemAllFields/roleassignments/addroleassignment(principalid=$userId,roleDefId=$roleDefId)"
        Invoke-SPRestWithThrottleHandling -Uri $assignUri -Method "POST" -Headers $spHeaders | Out-Null
        Write-Host "  ✓ Permission granted via SharePoint REST API!" -ForegroundColor Green
        Write-Host "    User: $Email" -ForegroundColor Cyan
        Write-Host "    Role: $permissionLevel" -ForegroundColor Cyan
        Write-Host "    Method: SharePoint REST (direct role assignment)" -ForegroundColor Cyan
        return @{ Success = $true; Method = "SharePoint-REST"; PermissionLevel = $permissionLevel }
    }
    catch {
        Write-Host "  ❌ Failed to assign role: $($_.Exception.Message)" -ForegroundColor Red
        if ($debug) {
            Write-Host "    Full error: $($_.Exception)" -ForegroundColor Gray
        }
        return $null
    }
}
#endregion Grant-DriveItemPermissionViaREST
#endregion REST Fallback Functions

#region Main Execution
#############################################################
#                    MAIN SCRIPT EXECUTION                  #
#############################################################

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  SharePoint/OneDrive File Permission" -ForegroundColor Cyan
Write-Host "  Grant Tool using Microsoft Graph API" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Authenticate with Microsoft Graph
AcquireToken

# Determine the site ID and item ID
$resolvedSiteId = $null
$resolvedDriveId = $null
$resolvedItemId = $null

if (-not [string]::IsNullOrWhiteSpace($fileLink)) {
    # Parse the sharing link to get site and item information
    Write-Host "`nResolving file from sharing link..." -ForegroundColor Cyan
    $itemInfo = Get-DriveItemFromLink -SharingLink $fileLink
    
    if ($itemInfo) {
        $resolvedSiteId = $itemInfo.SiteId
        $resolvedDriveId = $itemInfo.DriveId
        $resolvedItemId = $itemInfo.ItemId
    }
    else {
        Write-Host "❌ Could not resolve the sharing link. Please check the URL and try again." -ForegroundColor Red
        Exit
    }
}
elseif (-not [string]::IsNullOrWhiteSpace($siteId) -and -not [string]::IsNullOrWhiteSpace($itemId)) {
    # Use the directly specified site ID and item ID
    Write-Host "`nUsing directly specified site ID and item ID..." -ForegroundColor Cyan
    $resolvedSiteId = $siteId
    $resolvedItemId = $itemId
    
    # We need to get the drive ID from the site
    try {
        Test-ValidToken
        $headers = @{ "Authorization" = "Bearer $global:token" }
        $driveUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive"
        $driveResponse = Invoke-GraphRequestWithThrottleHandling -Uri $driveUri -Method "GET" -Headers $headers
        $resolvedDriveId = $driveResponse.id
        Write-Host "  Drive ID: $resolvedDriveId" -ForegroundColor Cyan
    }
    catch {
        Write-Host "❌ Could not retrieve drive information for the specified site." -ForegroundColor Red
        Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Yellow
        Exit
    }
}
else {
    Write-Host "❌ No file specified. Please provide either:" -ForegroundColor Red
    Write-Host "   - A direct link to the file (fileLink variable)" -ForegroundColor Yellow
    Write-Host "   - Both site ID and item ID (siteId and itemId variables)" -ForegroundColor Yellow
    Exit
}

# Validate required parameters
if ([string]::IsNullOrWhiteSpace($email)) {
    Write-Host "❌ No recipient email specified. Please set the `$email variable." -ForegroundColor Red
    Exit
}

# Grant permissions to the user via Graph API
$result = Grant-DriveItemPermission `
    -SiteId $resolvedSiteId `
    -DriveId $resolvedDriveId `
    -ItemId $resolvedItemId `
    -Email $email `
    -Message $message `
    -RequireSignIn $requireSignIn `
    -SendInvitation $sendInvitation `
    -Roles $roles

# If Graph API failed (common on IB-enabled sites), fall back to SharePoint REST API
if (-not $result -and $useRESTFallback) {
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "  Graph API failed - trying REST fallback" -ForegroundColor Yellow
    Write-Host "  (This is expected for IB-enabled sites)" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow

    # Determine the file URL and server-relative path for the REST fallback
    $restFileUrl = $fileLink  # Always pass the original link as context
    $restServerRelativeUrl = $null
    $restSiteUrl = $null

    if ($itemInfo) {
        # Prefer the actual server-relative path built from Graph drive webUrl + parentReference
        if ($itemInfo.ServerRelativeUrl) {
            $restServerRelativeUrl = $itemInfo.ServerRelativeUrl
            if ($debug) {
                Write-Host "  Using server-relative path from Graph: $restServerRelativeUrl" -ForegroundColor Gray
            }
        }
        # Use the resolved site URL from Graph
        if ($itemInfo.SiteUrl) {
            $restSiteUrl = $itemInfo.SiteUrl
        }
        # Use webUrl as the file URL context
        if ($itemInfo.DriveItem -and $itemInfo.DriveItem.webUrl) {
            $restFileUrl = $itemInfo.DriveItem.webUrl
        }
    }

    # Additional fallback: if ServerRelativeUrl is still null, try parsing the original $fileLink
    # Direct file URLs like https://tenant.sharepoint.com/sites/site/Shared%20Documents/file.docx
    # contain the correct server-relative path
    if (-not $restServerRelativeUrl -and -not [string]::IsNullOrWhiteSpace($fileLink)) {
        $linkUri = [System.Uri]$fileLink
        $linkPath = [System.Web.HttpUtility]::UrlDecode($linkUri.AbsolutePath)
        # Only use if it's a direct file path (not _layouts, not a sharing link)
        if ($linkPath -notmatch "_layouts/|_api/|/_vti_" -and $linkPath -match "\.[a-zA-Z0-9]+$") {
            $restServerRelativeUrl = $linkPath
            if ($debug) {
                Write-Host "  Using server-relative path from original fileLink: $restServerRelativeUrl" -ForegroundColor Gray
            }
        }
    }

    if ($restFileUrl) {
        $restParams = @{
            FileUrl = $restFileUrl
            Email   = $email
            Roles   = $roles
        }
        if ($restServerRelativeUrl) { $restParams['ServerRelativeUrl'] = $restServerRelativeUrl }
        if ($restSiteUrl) { $restParams['SiteUrl'] = $restSiteUrl }

        $result = Grant-DriveItemPermissionViaREST @restParams
    }
    else {
        Write-Host "  ❌ Cannot determine file URL for REST fallback" -ForegroundColor Red
    }
}

if ($result) {
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "  Permission granted successfully!" -ForegroundColor Green
    if ($result -is [hashtable] -and $result.Method) {
        Write-Host "  Method: $($result.Method)" -ForegroundColor Green
    }
    Write-Host "========================================`n" -ForegroundColor Green
}
else {
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "  Failed to grant permission" -ForegroundColor Red
    Write-Host "  Both Graph API and SharePoint REST fallback failed" -ForegroundColor Red
    Write-Host "========================================`n" -ForegroundColor Red
}
#endregion Main Execution
