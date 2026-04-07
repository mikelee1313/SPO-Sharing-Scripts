<#
.SYNOPSIS
    Audits SharePoint Online permission changes across a list of site collections
    using the Microsoft 365 Management Activity API.

.DESCRIPTION
    Uses the Office 365 Management Activity API (REST) to retrieve Audit.SharePoint
    content blobs and filter for SPO permission events:
      - SharingPermissionChanged  (permissions added/changed on an item)
      - PermissionLevelAdded      (new permission level added to a site)
      - PermissionLevelChanged    (permission level modified)
      - AddedToGroup              (user added to a SharePoint group)
      - RemovedFromGroup          (user removed from a SharePoint group)
      - SiteCollectionAdminAdded  (site collection admin added)
      - SiteCollectionAdminRemoved(site collection admin removed)
      - UniquePermissionsSet      (unique permissions created / inheritance broken)
      - SharingLinkCreated        (sharing link created for an item)
      - AddedToSharingLink        (user added to an existing sharing link)
      - SecureLinkCreated         (specific-people link created)
      - SecureLinkUpdated         (specific-people link modified)
      - AddedToSecureLink         (user added to a specific-people link)
      - RemovedFromSecureLink     (user removed from a specific-people link)
      - SharingInheritanceBroken  (unique permissions set / inheritance broken)

    PREREQUISITES:
      1. An Azure AD app registration with the ActivityFeed.Read APPLICATION permission
         granted on the "Office 365 Management APIs" resource (manage.office.com).
      2. Admin consent granted for that permission in your tenant.
      3. Unified audit logging enabled for the tenant.

    NOTE: The Management Activity API returns Audit.SharePoint data for the entire
    tenant. Filtering to the requested site URLs is performed client-side after
    downloading each content blob. The API enforces a maximum 24-hour window per
    request; this script automatically chunks longer ranges into 24-hour slices.
    The API only retains content for up to 7 days, so StartDate cannot be more
    than 7 days in the past.

    API Reference:
    https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference

.PARAMETER TenantId
    Azure AD tenant ID (GUID). Also used as the PublisherIdentifier for throttling.

.PARAMETER ClientId
    App registration client ID (GUID).

.PARAMETER ClientSecret
    App registration client secret (plain string). Used when authenticating with a secret.
    Cannot be combined with -CertificateThumbprint or -CertificatePath.

.PARAMETER CertificateThumbprint
    Thumbprint of a certificate already installed in the current user's or local machine's
    certificate store (Cert:\CurrentUser\My or Cert:\LocalMachine\My).
    Cannot be combined with -ClientSecret or -CertificatePath.

.PARAMETER CertificatePath
    Path to a .pfx certificate file. Use -CertificatePassword if the file is password-protected.
    Cannot be combined with -ClientSecret or -CertificateThumbprint.

.PARAMETER CertificatePassword
    Password for the .pfx file specified in -CertificatePath (optional).

.PARAMETER SiteListPath
    Path to a text file containing one SPO site URL per line.

.PARAMETER StartDate
    Start of the audit window (UTC). Cannot be more than 7 days in the past.
    Defaults to 7 days ago.

.PARAMETER EndDate
    End of the audit window (UTC). Defaults to now.

.PARAMETER OutputPath
    Path for the exported CSV. Defaults to .\SPO-PermissionsAudit_<timestamp>.csv

.PARAMETER IncludeSystemEvents
    When specified, includes internal SPO system-group rows (auto-generated
    Limited Access side-effects). Suppressed by default.

.EXAMPLE
    # Run with hardcoded secret defaults — no prompts
    .\SPO-permissionsAudit-MgmtAPI.ps1 -SiteListPath .\sites.txt

.EXAMPLE
    # Authenticate with a client secret
    .\SPO-permissionsAudit-MgmtAPI.ps1 `
        -TenantId     'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' `
        -ClientId     'yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy' `
        -ClientSecret 'your-client-secret' `
        -SiteListPath .\sites.txt

.EXAMPLE
    # Authenticate with a certificate thumbprint from the local cert store
    .\SPO-permissionsAudit-MgmtAPI.ps1 `
        -TenantId              'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' `
        -ClientId              'yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy' `
        -CertificateThumbprint 'A1B2C3D4E5F6...' `
        -SiteListPath          .\sites.txt

.EXAMPLE
    # Authenticate with a .pfx certificate file
    $pfxPass = Read-Host 'PFX password' -AsSecureString
    .\SPO-permissionsAudit-MgmtAPI.ps1 `
        -TenantId           'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' `
        -ClientId           'yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy' `
        -CertificatePath    'C:\certs\myapp.pfx' `
        -CertificatePassword $pfxPass `
        -SiteListPath       .\sites.txt

.NOTES
    Author  : Mike Lee / Mariel Williams
    Created : 3/27/2026
    Updated : 4/7/2026 - Converted from Search-UnifiedAuditLog to Management Activity API
              4/7/2026 - Added certificate-based authentication (thumbprint or .pfx file)
    Version : 2.1
#>

[CmdletBinding(DefaultParameterSetName = 'Secret')]
param (
    [Parameter()]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$TenantId = "9cfc42cb-51da-4055-87e9-b20a170b6ba3",

    [Parameter()]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$ClientId = "1e488dc4-1977-48ef-8d4d-9856f4e04536",

    # ── Auth option 1: client secret ────────────────────────────────────────────
    [Parameter(ParameterSetName = 'Secret')]
    [string]$ClientSecret = "",

    # ── Auth option 2: certificate thumbprint (from cert store) ──────────────────
    [Parameter(Mandatory, ParameterSetName = 'CertThumbprint')]
    [string]$CertificateThumbprint = "16f5dd7327719bc8cf15ff3c077adf59ace0c23",

    # ── Auth option 3: certificate .pfx file ────────────────────────────────────
    [Parameter(Mandatory, ParameterSetName = 'CertFile')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CertificatePath,

    [Parameter(ParameterSetName = 'CertFile')]
    [System.Security.SecureString]$CertificatePassword,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$SiteListPath = "C:\temp\SPOSiteList.txt",

    [Parameter()]
    [ValidateScript({
            if ($_ -lt (Get-Date).ToUniversalTime().AddDays(-7.5)) {
                throw "StartDate cannot be more than 7 days in the past (Management Activity API limitation)."
            }
            $true
        })]
    [datetime]$StartDate = (Get-Date).ToUniversalTime().AddDays(-7),

    [Parameter()]
    [datetime]$EndDate = (Get-Date).ToUniversalTime(),

    [Parameter()]
    [string]$OutputPath = ".\SPO-PermissionsAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter()]
    [switch]$IncludeSystemEvents
)

#region ── Configuration ────────────────────────────────────────────────────────

# Constructed after TenantId param is bound
$ApiBaseUrl = "https://manage.office.com/api/v1.0/$TenantId/activity/feed"
$ContentType = 'Audit.SharePoint'

# SPO permission-related operations to retain (all other operations are discarded client-side)
$PermissionOperations = @(
    'SharingPermissionChanged',
    'PermissionLevelAdded',
    'PermissionLevelChanged',
    'AddedToGroup',
    'RemovedFromGroup',
    'SiteCollectionAdminAdded',
    'SiteCollectionAdminRemoved',
    'UniquePermissionsSet',
    'SharingInheritanceBroken',
    'SharingSet',
    'AnonymousLinkCreated',
    'AnonymousLinkUpdated',
    'AnonymousLinkRemoved',
    'SharingLinkCreated',
    'AddedToSharingLink',
    'SecureLinkCreated',
    'SecureLinkUpdated',
    'AddedToSecureLink',
    'RemovedFromSecureLink'
)

#endregion

#region ── Helpers ──────────────────────────────────────────────────────────────

# Token cache — shared across all API calls; refreshed automatically before expiry
$script:AccessToken = $null
$script:TokenExpiry = [datetime]::MinValue

function New-ClientAssertionJwt {
    <#
    .SYNOPSIS  Builds and signs the client_assertion JWT required for certificate auth.
               The JWT is signed with the certificate's private key (RS256).
    #>
    param ([System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    # Header — convert hex thumbprint to bytes (compatible with PS 5.1 and PS 7)
    $thumbHex = $Certificate.Thumbprint -replace ' ', ''
    $thumbprintBytes = [byte[]]( 0..($thumbHex.Length / 2 - 1) | ForEach-Object {
            [Convert]::ToByte($thumbHex.Substring($_ * 2, 2), 16)
        })
    $x5t = [Convert]::ToBase64String($thumbprintBytes) -replace '\+', '-' -replace '/', '_' -replace '='
    $header = [Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{ alg = 'RS256'; typ = 'JWT'; x5t = $x5t } -Compress)
        )
    ) -replace '\+', '-' -replace '/', '_' -replace '='

    # Payload
    $now = [DateTimeOffset]::UtcNow
    $payload = [Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{
                aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
                iss = $ClientId
                sub = $ClientId
                jti = [Guid]::NewGuid().ToString()
                nbf = $now.ToUnixTimeSeconds()
                exp = $now.AddMinutes(10).ToUnixTimeSeconds()
            } -Compress)
        )
    ) -replace '\+', '-' -replace '/', '_' -replace '='

    # Signature (RS256)
    # GetRSAPrivateKey() is a .NET extension method — call it explicitly so PS 5.1 can resolve it.
    # Fall back to the legacy .PrivateKey property (RSACryptoServiceProvider) if needed.
    $rsa = $null
    try {
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    }
    catch { }

    if (-not $rsa) {
        # Legacy fallback: .PrivateKey returns RSACryptoServiceProvider on PS 5.1
        $rsa = $Certificate.PrivateKey
    }
    if (-not $rsa) {
        throw "Certificate '$($Certificate.Thumbprint)' does not have an accessible RSA private key. " +
        "Ensure the certificate was loaded with its private key and that the current user has permission to access it."
    }

    $sigInput = [System.Text.Encoding]::ASCII.GetBytes("$header.$payload")

    # RSACryptoServiceProvider (PS 5.1 legacy) uses a different SignData overload
    if ($rsa -is [System.Security.Cryptography.RSACryptoServiceProvider]) {
        $sigBytes = $rsa.SignData($sigInput, [System.Security.Cryptography.SHA256]::Create())
    }
    else {
        $sigBytes = $rsa.SignData($sigInput, [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    }
    $sig = [Convert]::ToBase64String($sigBytes) -replace '\+', '-' -replace '/', '_' -replace '='

    return "$header.$payload.$sig"
}

function Get-AccessToken {
    <#
    .SYNOPSIS  Returns a cached Bearer token, acquiring a new one when near expiry.
               Supports client secret (default) or certificate (thumbprint or .pfx file).
    #>
    if ($script:AccessToken -and (Get-Date).ToUniversalTime() -lt $script:TokenExpiry.AddMinutes(-5)) {
        return $script:AccessToken
    }

    Write-Verbose "Acquiring access token from Microsoft Entra ID..."

    # Build the token request body based on which auth method was supplied
    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scope = 'https://manage.office.com/.default'

    # Detect auth method from which script-level parameter was populated.
    # NOTE: $PSCmdlet.ParameterSetName inside a nested function refers to the
    # function's own binding, not the script's — so we check the variables directly.
    $useCert = $CertificateThumbprint -or $CertificatePath

    if ($useCert) {
        # ── Certificate auth: build client_assertion JWT ────────────────────────
        if ($CertificateThumbprint) {
            # Search CurrentUser first, then LocalMachine
            $cert = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
            if (-not $cert) {
                $cert = Get-Item "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
            }
            if (-not $cert) {
                throw "Certificate with thumbprint '$CertificateThumbprint' not found in CurrentUser\My or LocalMachine\My."
            }
        }
        else {
            # Load from .pfx file
            $cert = if ($CertificatePassword) {
                [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                    (Resolve-Path $CertificatePath).Path,
                    $CertificatePassword,
                    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
                )
            }
            else {
                [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                    (Resolve-Path $CertificatePath).Path,
                    [string]$null,
                    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
                )
            }
        }

        $assertion = New-ClientAssertionJwt -Certificate $cert
        $body = "grant_type=client_credentials" +
        "&client_id=$([Uri]::EscapeDataString($ClientId))" +
        "&client_assertion_type=$([Uri]::EscapeDataString('urn:ietf:params:oauth:client-assertion-type:jwt-bearer'))" +
        "&client_assertion=$([Uri]::EscapeDataString($assertion))" +
        "&scope=$([Uri]::EscapeDataString($scope))"
    }
    else {
        # ── Client secret auth ───────────────────────────────────────────────────
        $body = "grant_type=client_credentials" +
        "&client_id=$([Uri]::EscapeDataString($ClientId))" +
        "&client_secret=$([Uri]::EscapeDataString($ClientSecret))" +
        "&scope=$([Uri]::EscapeDataString($scope))"
    }

    try {
        $resp = Invoke-RestMethod `
            -Method      Post `
            -Uri         $tokenUri `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body        $body `
            -ErrorAction Stop

        $script:AccessToken = $resp.access_token
        $script:TokenExpiry = (Get-Date).ToUniversalTime().AddSeconds([int]$resp.expires_in)
        Write-Verbose "  Token acquired; expires at $($script:TokenExpiry.ToString('u'))"
        return $script:AccessToken
    }
    catch {
        throw "Failed to acquire access token: $_"
    }
}

function Invoke-MgmtApi {
    <#
    .SYNOPSIS  Calls a Management Activity API endpoint with retry on throttling/transient errors.
    #>
    param (
        [string]$Uri,
        [string]$Method = 'GET',
        [string]$Body = $null,
        [int]   $MaxRetries = 5
    )

    $attempt = 0
    while ($true) {
        $attempt++
        $token = Get-AccessToken

        $params = @{
            Uri             = $Uri
            Method          = $Method
            Headers         = @{ Authorization = "Bearer $token" }
            UseBasicParsing = $true
            ErrorAction     = 'Stop'
        }
        if ($Body) {
            $params.ContentType = 'application/json; charset=utf-8'
            $params.Body = $Body
        }

        try {
            return Invoke-WebRequest @params
        }
        catch {
            # Inspect HTTP status code — works for both PS 5.1 (WebException) and PS 7 (HttpResponseException)
            $statusCode = 0
            if ($null -ne $_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($statusCode -eq 429 -and $attempt -lt $MaxRetries) {
                $retryAfter = $_.Exception.Response.Headers['Retry-After']
                $delay = if ($retryAfter) { [int]$retryAfter } else { [math]::Pow(2, $attempt) * 5 }
                Write-Warning "    Rate limited (429). Retrying in ${delay}s (attempt $attempt/$MaxRetries)..."
                Start-Sleep -Seconds $delay
            }
            elseif ($statusCode -in 500, 503 -and $attempt -lt $MaxRetries) {
                $delay = [math]::Pow(2, $attempt) * 3
                Write-Warning "    Server error ($statusCode). Retrying in ${delay}s (attempt $attempt/$MaxRetries)..."
                Start-Sleep -Seconds $delay
            }
            else {
                throw
            }
        }
    }
}

function Assert-AuditSubscription {
    <#
    .SYNOPSIS  Ensures the Audit.SharePoint subscription is active. Creates it if absent.
               Calling /subscriptions/start on an existing subscription is a safe no-op.
    #>
    Write-Verbose "Ensuring Audit.SharePoint subscription is active..."
    $uri = "$ApiBaseUrl/subscriptions/start?contentType=$ContentType&PublisherIdentifier=$TenantId"

    try {
        $resp = Invoke-MgmtApi -Uri $uri -Method 'POST'
        $sub = $resp.Content | ConvertFrom-Json
        Write-Verbose "  Subscription status: $($sub.status)"

        if ($sub.status -ne 'enabled') {
            throw "Subscription returned status '$($sub.status)' — verify app permissions and admin consent."
        }
    }
    catch {
        # Fall back: check whether an enabled subscription already exists
        Write-Verbose "  /subscriptions/start returned an error; verifying via /subscriptions/list..."
        try {
            $listResp = Invoke-MgmtApi -Uri "$ApiBaseUrl/subscriptions/list?PublisherIdentifier=$TenantId"
            $existing = ($listResp.Content | ConvertFrom-Json) |
            Where-Object { $_.contentType -eq $ContentType -and $_.status -eq 'enabled' }

            if (-not $existing) {
                throw "No active Audit.SharePoint subscription found. Original error: $_"
            }
            Write-Verbose "  Existing enabled subscription confirmed."
        }
        catch {
            throw "Failed to start or verify Audit.SharePoint subscription: $_"
        }
    }
}

function Get-ContentBlobList {
    <#
    .SYNOPSIS  Lists all available content blobs for a single ≤24-hour window,
               following NextPageUri pagination until exhausted.
    #>
    param (
        [datetime]$Start,
        [datetime]$End
    )

    $blobs = [System.Collections.Generic.List[PSObject]]::new()
    $startStr = $Start.ToString('yyyy-MM-ddTHH:mm:ss')
    $endStr = $End.ToString('yyyy-MM-ddTHH:mm:ss')
    $uri = "$ApiBaseUrl/subscriptions/content?contentType=$ContentType" +
    "&startTime=$startStr&endTime=$endStr&PublisherIdentifier=$TenantId"

    do {
        Write-Verbose "    Listing blobs: $uri"
        $response = Invoke-MgmtApi -Uri $uri
        $page = $response.Content | ConvertFrom-Json
        if ($page) { foreach ($blob in $page) { $blobs.Add($blob) } }

        # NextPageUri header signals additional pages (handle both PS 5.1 NameValueCollection and PS 7 Dictionary)
        $nextHeader = $response.Headers['NextPageUri']
        $uri = if ($nextHeader) {
            if ($nextHeader -is [System.Array]) { $nextHeader[0] } else { $nextHeader }
        }
        else { $null }

    } while ($uri)

    return $blobs
}

function Get-BlobEvents {
    <#
    .SYNOPSIS  Downloads a content blob by its URI and returns the array of audit events.
    #>
    param ([string]$ContentUri)

    # PublisherIdentifier must be appended to the content URI as well
    $uri = "${ContentUri}?PublisherIdentifier=$TenantId"
    $response = Invoke-MgmtApi -Uri $uri
    return $response.Content | ConvertFrom-Json
}

# Friendly display names for each audit operation
$script:ActionLabels = @{
    'SharingInheritanceBroken'   = 'Unique Permissions Created (Inheritance Broken)'
    'UniquePermissionsSet'       = 'Unique Permissions Set'
    'AddedToGroup'               = 'User Added to Group'
    'RemovedFromGroup'           = 'User Removed from Group'
    'SharingSet'                 = 'Permissions Granted'
    'SharingPermissionChanged'   = 'Permission Changed'
    'SharingLinkCreated'         = 'Sharing Link Created'
    'AddedToSharingLink'         = 'User Added to Sharing Link'
    'SecureLinkCreated'          = 'Secure Link Created (Specific People)'
    'SecureLinkUpdated'          = 'Secure Link Updated (Specific People)'
    'AddedToSecureLink'          = 'User Added to Secure Link'
    'RemovedFromSecureLink'      = 'User Removed from Secure Link'
    'AnonymousLinkCreated'       = 'Anonymous Link Created'
    'AnonymousLinkUpdated'       = 'Anonymous Link Updated'
    'AnonymousLinkRemoved'       = 'Anonymous Link Removed'
    'PermissionLevelAdded'       = 'Permission Level Added'
    'PermissionLevelChanged'     = 'Permission Level Modified'
    'SiteCollectionAdminAdded'   = 'Site Collection Admin Added'
    'SiteCollectionAdminRemoved' = 'Site Collection Admin Removed'
}

function ConvertTo-FlatRecord {
    <#
    .SYNOPSIS  Flattens a Management Activity API audit event into a clean, admin-readable object.

    NOTE: Unlike the UAL cmdlet (which wraps data in an AuditData JSON string), the
    Management Activity API returns events as already-parsed JSON objects. Fields like
    SiteUrl, ObjectId, ClientIP, etc. are direct properties — no ConvertFrom-Json needed.
    #>
    param ([PSObject]$Event)

    $op = $Event.Operation
    $action = if ($script:ActionLabels[$op]) { $script:ActionLabels[$op] } else { $op }

    # Parse PermissionsGranted and GroupAffected out of the EventData XML blob (same as before)
    $permGranted = ''
    $groupName = ''
    if ($Event.EventData) {
        if ($Event.EventData -match '<PermissionsGranted>([^<]+)<') { $permGranted = $Matches[1] }
        if ($Event.EventData -match '<Group>([^<]+)<') { $groupName = $Matches[1] }
    }

    # Relative path: SourceRelativeUrl is cleanest; fall back to stripping the site URL from ObjectId
    $relPath = ''
    if ($Event.SourceRelativeUrl) {
        $relPath = $Event.SourceRelativeUrl
    }
    elseif ($Event.ObjectId -and $Event.SiteUrl) {
        $stripped = $Event.ObjectId -replace [regex]::Escape($Event.SiteUrl.TrimEnd('/')), ''
        $relPath = if ($stripped -match '^[/\\]?$') { '(site root)' } else { $stripped.TrimStart('/') }
    }

    # Clean up target name — internal SharingLinks group names are not meaningful to admins
    $target = if ($Event.TargetUserOrGroupName -match '^SharingLinks\.') {
        '(sharing link group)'
    }
    else {
        $Event.TargetUserOrGroupName
    }

    # Flag system-generated side-effect rows so they can be filtered
    $isSystem = ($op -eq 'AddedToGroup') -and (
        $groupName -match '^Limited Access System Group' -or
        $groupName -match '^SharingLinks\.'
    )

    # Friendly link scope (blank when not applicable)
    $linkScope = if ($Event.SharingLinkScope -and $Event.SharingLinkScope -notin 'Uninitialized', 'None') {
        $Event.SharingLinkScope
    }
    else { '' }

    # Field mapping vs. UAL cmdlet:
    #   $Record.CreationDate  → $Event.CreationTime   (direct property)
    #   $Record.UserIds       → $Event.UserId         (direct property)
    #   $Record.Operations    → $Event.Operation      (direct property)
    #   $Record.AuditData     → $Event itself         (already parsed; no nested JSON)
    [PSCustomObject]@{
        DateTime           = $Event.CreationTime
        PerformedBy        = $Event.UserId
        Action             = $action
        ItemType           = $Event.ItemType
        RelativePath       = $relPath
        SiteUrl            = $Event.SiteUrl
        AffectedUser       = $target
        PermissionsGranted = $permGranted
        GroupAffected      = $groupName
        LinkScope          = $linkScope
        ClientIP           = $Event.ClientIP
        IsSystemEvent      = $isSystem
    }
}

#endregion

#region ── Main ─────────────────────────────────────────────────────────────────

# Load site list - skip blank lines and comment lines
$sites = Get-Content -Path $SiteListPath |
Where-Object { $_ -match 'https?://' } |
ForEach-Object { $_.Trim().TrimEnd('/') } |
Select-Object -Unique

if (-not $sites) {
    throw "No valid SPO URLs found in '$SiteListPath'. Each line should contain a URL starting with https://"
}

# HashSet for O(1) site URL lookups during client-side filtering
$sitesSet = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]$sites,
    [StringComparer]::OrdinalIgnoreCase
)

# HashSet for O(1) operation lookups when filtering blobs
$operationsSet = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]$PermissionOperations,
    [StringComparer]::OrdinalIgnoreCase
)

Write-Host "SPO Permissions Audit (Management Activity API)" -ForegroundColor Cyan
Write-Host "  Sites      : $($sites.Count)" -ForegroundColor Cyan
Write-Host "  Window     : $($StartDate.ToString('u'))  →  $($EndDate.ToString('u'))" -ForegroundColor Cyan
Write-Host "  Operations : $($PermissionOperations.Count) event types" -ForegroundColor Cyan
Write-Host ""

# Step 1 – Acquire token and confirm it contains the required permission
$authMethod = if ($CertificateThumbprint) { "Certificate (thumbprint: $CertificateThumbprint)" }
elseif ($CertificatePath) { "Certificate (file: $CertificatePath)" }
else { 'Client Secret' }
Write-Host "  Auth method: $authMethod" -ForegroundColor Cyan
Write-Host "  Acquiring access token..." -ForegroundColor Cyan
$testToken = Get-AccessToken
if (-not $testToken) {
    throw "Failed to acquire access token. Check TenantId, ClientId, and ClientSecret."
}

# Decode JWT payload (middle segment) to inspect the 'roles' claim without a module dependency
$jwtPayload = $testToken.Split('.')[1]
# Pad Base64 to a multiple of 4
$jwtPayload += '=' * ((4 - $jwtPayload.Length % 4) % 4)
$claims = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($jwtPayload)) | ConvertFrom-Json

$tokenRoles = @($claims.roles)
Write-Host "  Token roles : $($tokenRoles -join ', ')" -ForegroundColor Cyan

if ('ActivityFeed.Read' -notin $tokenRoles) {
    Write-Host ""
    Write-Host "  *** PERMISSION MISSING ***" -ForegroundColor Red
    Write-Host "  The access token does NOT contain 'ActivityFeed.Read'." -ForegroundColor Red
    Write-Host "  Fix in Azure Portal:" -ForegroundColor Yellow
    Write-Host "    1. Azure Active Directory → App registrations → your app" -ForegroundColor Yellow
    Write-Host "    2. API permissions → Add a permission" -ForegroundColor Yellow
    Write-Host "    3. APIs my organization uses → 'Office 365 Management APIs'" -ForegroundColor Yellow
    Write-Host "    4. Application permissions → check 'ActivityFeed.Read'" -ForegroundColor Yellow
    Write-Host "    5. Click 'Grant admin consent for <tenant>'" -ForegroundColor Yellow
    throw "ActivityFeed.Read application permission is not consented. See instructions above."
}
Write-Host "  Token OK — ActivityFeed.Read permission confirmed." -ForegroundColor Green
Write-Host ""

# Step 2 – Ensure the Audit.SharePoint subscription is running
Assert-AuditSubscription

# Step 3 – Split the date range into ≤24-hour chunks (API maximum per request)
#           Content blobs are retrieved by chunk; per-event filtering is client-side.
$chunks = [System.Collections.Generic.List[hashtable]]::new()
$cursor = $StartDate
while ($cursor -lt $EndDate) {
    $chunkEnd = $cursor.AddHours(24)
    if ($chunkEnd -gt $EndDate) { $chunkEnd = $EndDate }
    # Skip zero-duration or sub-second chunks (can occur due to millisecond differences between param defaults)
    if (($chunkEnd - $cursor).TotalSeconds -lt 1) { break }
    $chunks.Add(@{ Start = $cursor; End = $chunkEnd })
    $cursor = $chunkEnd
}

Write-Host "  Date chunks: $($chunks.Count) x ≤24-hour window(s)" -ForegroundColor Cyan
Write-Host ""

$totalWritten = 0
$totalFiltered = 0
$totalBlobs = 0
$chunkIndex = 0

foreach ($chunk in $chunks) {
    $chunkIndex++

    Write-Progress -Activity "Fetching Management Activity API blobs" `
        -Status       "Chunk $chunkIndex/$($chunks.Count): $($chunk.Start.ToString('u')) → $($chunk.End.ToString('u'))" `
        -PercentComplete (($chunkIndex / $chunks.Count) * 100)

    try {
        # Step 3 – List available content blobs for this 24-hour window
        $blobs = Get-ContentBlobList -Start $chunk.Start -End $chunk.End
        $totalBlobs += $blobs.Count
        Write-Verbose "  Chunk $chunkIndex — $($blobs.Count) content blob(s) available"

        foreach ($blob in $blobs) {
            Write-Verbose "    Downloading blob: $($blob.contentId)"
            try {
                # Step 4 – Download the blob (array of audit event objects)
                $events = Get-BlobEvents -ContentUri $blob.contentUri

                foreach ($event in $events) {
                    # Filter 1: keep only the permission-related operations
                    if (-not $operationsSet.Contains($event.Operation)) { continue }

                    # Filter 2: keep only events that belong to one of the requested sites
                    $normalizedSite = if ($event.SiteUrl) { $event.SiteUrl.TrimEnd('/') } else { '' }
                    if ($normalizedSite -and -not $sitesSet.Contains($normalizedSite)) { continue }

                    $flat = ConvertTo-FlatRecord -Event $event

                    if ($flat.IsSystemEvent -and -not $IncludeSystemEvents) {
                        $totalFiltered++
                    }
                    else {
                        $flat |
                        Select-Object DateTime, PerformedBy, Action, ItemType, RelativePath, SiteUrl,
                        AffectedUser, PermissionsGranted, GroupAffected, LinkScope, ClientIP |
                        Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Append
                        $totalWritten++
                    }
                }
            }
            catch {
                Write-Warning "  [!] Error fetching blob '$($blob.contentId)': $_"
            }
        }
    }
    catch {
        Write-Warning "  [!] Error listing content for chunk $chunkIndex ($($chunk.Start.ToString('u'))): $_"
    }
}

Write-Progress -Activity "Fetching Management Activity API blobs" -Completed

#endregion

#region ── Output ───────────────────────────────────────────────────────────────

Write-Host "  Content blobs processed: $totalBlobs" -ForegroundColor Cyan

if ($totalWritten -eq 0) {
    Write-Host "`nNo permission events found across any sites in the specified window." -ForegroundColor Yellow
}
else {
    Write-Host "`nResults  : $totalWritten permission change events" -ForegroundColor Cyan
    if (-not $IncludeSystemEvents -and $totalFiltered -gt 0) {
        Write-Host "Filtered : $totalFiltered internal SPO system-group events suppressed (use -IncludeSystemEvents to include)" -ForegroundColor DarkGray
    }
    Write-Host "Exported : $OutputPath" -ForegroundColor Green

    # Read back the CSV for summary reporting (avoids holding all records in memory)
    $exportData = Import-Csv -Path $OutputPath

    # Summary by action type
    Write-Host "`n── Events by action ────────────────────────────────────" -ForegroundColor Cyan
    $exportData |
    Group-Object Action |
    Sort-Object Count -Descending |
    Format-Table @{L = 'Action'; E = { $_.Name }; W = 50 }, Count -AutoSize

    # Summary by site
    Write-Host "── Events per site ─────────────────────────────────────" -ForegroundColor Cyan
    $exportData |
    Group-Object SiteUrl |
    Sort-Object Count -Descending |
    Format-Table @{L = 'SiteUrl'; E = { $_.Name }; W = 60 }, Count -AutoSize

    # Who performed changes
    Write-Host "── Changes by user ─────────────────────────────────────" -ForegroundColor Cyan
    $exportData |
    Group-Object PerformedBy |
    Sort-Object Count -Descending |
    Format-Table @{L = 'PerformedBy'; E = { $_.Name }; W = 50 }, Count -AutoSize
}

#endregion
