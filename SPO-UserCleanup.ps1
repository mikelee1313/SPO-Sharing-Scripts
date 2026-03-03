<#
.SYNOPSIS
    SharePoint Online Permissions Remover - Report and Remove specified users from SPO sites based on direct permissions, group memberships, and other access vectors.
    This script helps will identify what sites a user has access to and then remove them from those sites.
    This can be used for cleanup of former employees that may be rehired in the future, then generate a PUID mismatch.

.DESCRIPTION

    REPORT  - Scans all (or a target) SPO site(s) for specified users and writes
              a CSV audit report: SiteName, URL, User, Owner, AccessType.

    REMOVE  - Reads a previously generated report CSV and removes each user from
              only the site(s) where they were found (targeted removal).

    BOTH    - Runs REPORT first, then automatically feeds the output CSV into
              the REMOVE phase in a single execution.

    The REPORT checks every access vector:
      * Direct user permissions (with permission-level detail)
      * Microsoft 365 Group membership (member & owner)
      * Site ownership
      * Site Collection Administration
      * SharePoint Group membership
      * Entra ID / Azure AD security group membership
      * "Everyone except external users" (optional)
      * User Information List presence (always reported; EEEU activity noted in AccessType to help distinguish residual vs phantom entries)

    The REMOVE phase targets:
      * SharePoint group memberships
      * Direct file/item permissions
      * Sharing link group memberships
      * User Information List (optional)

.PARAMETER Mode
    "Report"  - Audit only; produces CSV.
    "Remove"  - Remove users listed in an existing CSV.
    "Both"    - Audit then remove in one pass.

.NOTES
    File Name      : SPO-UserCleanup.ps1
    Authors        : Mike Lee 
    Created        : 3/3/2026

        - PnP PowerShell module  (Install-Module PnP.PowerShell)
        - Entra ID App Registration with:
            SharePoint  ->  Sites.FullControl.All  (Application)
            Graph       ->  Sites.Read.All          (Application)
        - Certificate-based App-Only authentication

.EXAMPLE
    # Audit all sites, produce CSV
    $Mode = "Report"; .\SPO-UserCleanup.ps1

    # Remove using an existing CSV
    $Mode = "Remove"; $InputCsvPath = "C:\temp\SiteUsers_2026-03-02_output.csv"
    .\SPO-UserCleanup.ps1

    # Audit then immediately remove
    $Mode = "Both"; .\SPO-UserCleanup.ps1
#>

#=================================================================================================
# USER CONFIGURATION  -  Update ALL variables in this section before running
#=================================================================================================

# --- App-Only Authentication (same app registration used for both modes) ---
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"    # Entra (Azure AD) App/Client ID
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082" # Certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"    # Tenant ID (GUID)
$t = 'M365CPI13246019'                          # Tenant name (no .onmicrosoft.com)

# --- Mode Selection ---
#   "Report"  – scan sites, write CSV
#   "Remove"  – read CSV produced by Report mode, remove users from found sites
#   "Both"    – run Report then Remove in one execution
$Mode = "Report"

# ---- REPORT mode settings -----------------------------------------------------------------------
# Path to a text file containing user UPNs to search for (one per line)
$UsersFilePath = 'C:\temp\users.txt'

# Leave empty to process ALL sites; set to a specific URL to target one site
$TargetSiteUrl = ""

# Include OneDrive for Business sites in the scan
$IncludeOneDrive = $false

# Check "Everyone except external users" permissions (adds processing time)
$checkEEEU = $false

# ---- REMOVE mode settings -----------------------------------------------------------------------
# CSV produced by a prior Report run.  When Mode = "Both" this is set automatically.
$InputCsvPath = ""

# Remove users from the Site's User Information List (UIL) as well as permissions/groups
#(optional; may cause issues if user has residual UIL entry but no actual access)
$RemoveFromUIL = $true

# ---- Shared / output settings -------------------------------------------------------------------
$debug = $false  # $true for verbose console + log output

# Throttling protection (highly recommended for large tenants)
$enableThrottlingProtection = $true
$baseDelayBetweenSites = 2    # seconds between sites
$baseDelayBetweenUsers = 1    # seconds between users within a site
$maxRetryAttempts = 5    # max retry attempts on throttle
$baseRetryDelay = 30   # base retry delay (seconds); uses exponential backoff

#=================================================================================================
# END OF USER CONFIGURATION
#=================================================================================================

#region ── Initialization ─────────────────────────────────────────────────────────────────────────

$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$outputFile = "$env:TEMP\SiteUsers_$($date)_output.csv"   # Report-mode output CSV
$log = "$env:TEMP\SiteUsers_$($date)_logfile.log"  # Shared log file
$firstWrite = $true   # tracks whether CSV header has been written yet

# Validate Mode
if ($Mode -notin @("Report", "Remove", "Both")) {
    Write-Host "ERROR: `$Mode must be 'Report', 'Remove', or 'Both'." -ForegroundColor Red
    exit 1
}

#endregion ────────────────────────────────────────────────────────────────────────────────────────

#region ── Shared Helper Functions ────────────────────────────────────────────────────────────────

# Logging -----------------------------------------------------------------------------------------
function Write-LogEntry {
    param([string]$LogName, [string]$LogEntryText)
    if ($LogName) {
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : $LogEntryText" |
        Out-File -FilePath $LogName -Append
    }
}

function Write-DebugInfo {
    param([string]$Message, [string]$ForegroundColor = "Gray", [switch]$AlwaysLog)
    if ($debug) { Write-Host $Message -ForegroundColor $ForegroundColor }
    if ($debug -or $AlwaysLog) { Write-LogEntry -LogName $log -LogEntryText $Message }
}

function Write-InfoMessage {
    param([string]$Message, [string]$ForegroundColor = "White", [switch]$AlwaysShow)
    if ($debug -or $AlwaysShow) { Write-Host $Message -ForegroundColor $ForegroundColor }
    Write-LogEntry -LogName $log -LogEntryText $Message
}

function Write-StatusMessage {
    param([string]$Message, [string]$ForegroundColor = "Yellow", [switch]$Force)
    Write-Host $Message -ForegroundColor $ForegroundColor
    Write-LogEntry -LogName $log -LogEntryText $Message
}

# Throttling --------------------------------------------------------------------------------------
function Invoke-PnPCommandWithThrottling {
    param(
        [scriptblock]$Command,
        [string]$OperationDescription = "SharePoint operation",
        [int]$MaxRetries = $maxRetryAttempts,
        [int]$BaseDelay = $baseRetryDelay
    )

    if (-not $enableThrottlingProtection) { return & $Command }

    $attempt = 0
    $delay = $BaseDelay

    while ($attempt -lt $MaxRetries) {
        try {
            $attempt++
            Write-DebugInfo "   Executing: $OperationDescription (attempt $attempt/$MaxRetries)"
            $result = & $Command
            if ($attempt -gt 1) {
                Write-StatusMessage "    ✓ $OperationDescription succeeded after $attempt attempts" -ForegroundColor Green -Force
            }
            return $result
        }
        catch {
            $statusCode = $null
            $retryAfter = $null

            if ($_.Exception.Message -match "429|Too Many Requests") {
                $statusCode = 429
                Write-StatusMessage "    ⚠️ HTTP 429 (Too Many Requests) – $OperationDescription" -ForegroundColor Yellow -Force
            }
            elseif ($_.Exception.Message -match "503|Server Too Busy") {
                $statusCode = 503
                Write-StatusMessage "    ⚠️ HTTP 503 (Server Too Busy) – $OperationDescription" -ForegroundColor Yellow -Force
            }
            elseif ($_.Exception.Message -match "throttl|rate limit") {
                $statusCode = 429
                Write-StatusMessage "    ⚠️ Throttling detected – $OperationDescription" -ForegroundColor Yellow -Force
            }

            if (($statusCode -eq 429 -or $statusCode -eq 503) -and ($attempt -lt $MaxRetries)) {
                if ($_.Exception.Message -match "Retry-After[:\s]+(\d+)") { $retryAfter = [int]$matches[1] }
                $waitTime = if ($retryAfter -and $retryAfter -gt 0) {
                    $retryAfter
                }
                else {
                    ($delay * [Math]::Pow(2, $attempt - 1)) + (Get-Random -Minimum 0 -Maximum 10)
                }
                Write-StatusMessage "    ⏳ Waiting $waitTime seconds before retry $($attempt+1)/$MaxRetries..." -ForegroundColor Yellow -Force
                Write-LogEntry -LogName $log -LogEntryText "Throttling on $OperationDescription. Waiting $waitTime s (retry $($attempt+1)/$MaxRetries)"
                Start-Sleep -Seconds $waitTime
                continue
            }
            else {
                if ($attempt -ge $MaxRetries) {
                    Write-StatusMessage "    ❌ Max retries ($MaxRetries) exceeded: $OperationDescription" -ForegroundColor Red -Force
                    Write-LogEntry -LogName $log -LogEntryText "Max retries exceeded for $OperationDescription. Last error: $($_.Exception.Message)"
                }
                else {
                    Write-DebugInfo "    ❌ Non-throttling error: $OperationDescription – $($_.Exception.Message)" -ForegroundColor Red -AlwaysLog
                }
                throw
            }
        }
    }
}

function Add-ThrottlingDelay {
    param([string]$DelayType = "user", [string]$Description = "")
    if (-not $enableThrottlingProtection) { return }
    $delaySeconds = switch ($DelayType.ToLower()) {
        "site" { $baseDelayBetweenSites }
        "user" { $baseDelayBetweenUsers }
        default { 1 }
    }
    if ($delaySeconds -gt 0) {
        Write-DebugInfo "    ⏱️ Adding $delaySeconds s delay$Description"
        Start-Sleep -Seconds $delaySeconds
    }
}

function Invoke-WithThrottleHandling {
    # Alias used by Remove functions (matches SPOUserRemover style)
    param([scriptblock]$ScriptBlock, [string]$Operation, [int]$MaxRetries = 3, [int]$BaseDelaySeconds = 5)
    Invoke-PnPCommandWithThrottling -Command $ScriptBlock -OperationDescription $Operation `
        -MaxRetries $MaxRetries -BaseDelay $BaseDelaySeconds
}

# Entra ID group membership -----------------------------------------------------------------------
function Test-EntraGroupMembership {
    param([string]$UserPrincipalName, [string]$GroupId, [string]$GroupDisplayName)
    try {
        $groupMembers = Get-PnPMicrosoft365GroupMember -Identity $GroupId -ErrorAction SilentlyContinue
        if ($groupMembers) {
            $isMember = $groupMembers | Where-Object { $_.UserPrincipalName -eq $UserPrincipalName -or $_.Mail -eq $UserPrincipalName }
            if ($isMember) { return $true }
        }
        $groupOwners = Get-PnPMicrosoft365GroupOwner -Identity $GroupId -ErrorAction SilentlyContinue
        if ($groupOwners) {
            $isOwner = $groupOwners | Where-Object { $_.UserPrincipalName -eq $UserPrincipalName -or $_.Mail -eq $UserPrincipalName }
            if ($isOwner) {
                Write-DebugInfo "User is an OWNER of M365 group '$GroupDisplayName'" -ForegroundColor Magenta
                return $true
            }
        }
        $azureGroup = Get-PnPAzureADGroup -Identity $GroupId -ErrorAction SilentlyContinue
        if ($azureGroup) {
            $azureGroupMembers = Get-PnPAzureADGroupMember -Identity $GroupId -ErrorAction SilentlyContinue
            if ($azureGroupMembers) {
                $isMember = $azureGroupMembers | Where-Object { $_.UserPrincipalName -eq $UserPrincipalName -or $_.Mail -eq $UserPrincipalName }
                return ($null -ne $isMember)
            }
        }
        return $false
    }
    catch {
        Write-DebugInfo "Could not check membership for group '$GroupDisplayName': $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $false
    }
}

# Connect helper (site-level) ---------------------------------------------------------------------
function Connect-ToSite {
    param([string]$SiteUrl)
    Invoke-PnPCommandWithThrottling -Command {
        Connect-PnPOnline -Url $SiteUrl -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
    } -OperationDescription "Connect to site $SiteUrl" | Out-Null
}

#endregion ────────────────────────────────────────────────────────────────────────────────────────

#region ── REPORT Functions ───────────────────────────────────────────────────────────────────────

function Get-SiteList {
    if ($TargetSiteUrl -and $TargetSiteUrl -ne "") {
        Write-StatusMessage "Targeting single site: $TargetSiteUrl" -ForegroundColor Cyan
        Write-LogEntry -LogName $log -LogEntryText "Targeting single site: $TargetSiteUrl"
        try {
            $sites = @(Get-PnPTenantSite -Url $TargetSiteUrl -ErrorAction Stop)
            Write-StatusMessage "Retrieved target site: $($sites[0].Title)" -ForegroundColor Green
            return $sites
        }
        catch {
            Write-StatusMessage "Failed to retrieve target site '$TargetSiteUrl': $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    }

    Write-StatusMessage "Retrieving all sites from tenant..." -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Retrieving all sites from tenant..."

    $excludeTemplates = { $_.Template -ne 'RedirectSite#0' -and $_.Template -notlike 'SRCHCEN*' -and
        $_.Template -notlike 'SRCHCENTERLITE*' -and $_.Template -notlike 'SPSMSITEHOST*' -and
        $_.Template -notlike 'APPCATALOG*' -and $_.Template -notlike 'REDIRECTSITE*' }

    if ($IncludeOneDrive) {
        $sites = Get-PnPTenantSite -IncludeOneDriveSites | Where-Object $excludeTemplates
    }
    else {
        $sites = Get-PnPTenantSite | Where-Object $excludeTemplates
    }

    Write-StatusMessage "Retrieved $($sites.Count) sites to process" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Retrieved $($sites.Count) sites to process"
    return $sites
}

function Invoke-ReportMode {
    param([array]$Users, [switch]$RemoveInline)

    Write-StatusMessage "`n========================================" -ForegroundColor Cyan
    Write-StatusMessage "  REPORT MODE  -  Scanning for users" -ForegroundColor Cyan
    Write-StatusMessage "========================================`n" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting REPORT mode"

    $sites = Get-SiteList

    foreach ($site in $sites) {
        Add-ThrottlingDelay -DelayType "site" -Description " between sites"
        $siteOutput = @()

        Write-StatusMessage "Processing site: $($site.Title) ($($site.Url))" -ForegroundColor Yellow
        Write-LogEntry -LogName $log -LogEntryText "Starting processing: $($site.Title) ($($site.Url))"

        foreach ($user in $Users) {
            Add-ThrottlingDelay -DelayType "user" -Description " between users"

            Write-DebugInfo "Attempting to find '$user' on '$($site.Url)'" -ForegroundColor Green
            Write-LogEntry -LogName $log -LogEntryText "Checking '$user' on '$($site.Url)'"

            try {
                Connect-ToSite -SiteUrl $site.Url

                $userFound = $false
                $accessType = ""
                $groupMemberships = @()

                # ── Check 1: Direct permissions ─────────────────────────────────────────────
                try {
                    $usersWithRights = Invoke-PnPCommandWithThrottling -Command {
                        Get-PnPUser -WithRightsAssigned -ErrorAction SilentlyContinue
                    } -OperationDescription "Get users with rights assigned"

                    $userWithPerms = $usersWithRights | Where-Object {
                        $_.LoginName -eq $user -or $_.Email -eq $user -or
                        $_.UserPrincipalName -eq $user -or $_.LoginName -like "*$user*"
                    }

                    if ($userWithPerms) {
                        try {
                            $web = Invoke-PnPCommandWithThrottling -Command {
                                Get-PnPWeb -ErrorAction SilentlyContinue
                            } -OperationDescription "Get web for role assignments"

                            $roleAssignments = Invoke-PnPCommandWithThrottling -Command {
                                Get-PnPProperty -ClientObject $web -Property RoleAssignments -ErrorAction SilentlyContinue
                            } -OperationDescription "Get role assignments"

                            $permLevels = @()
                            foreach ($ra in $roleAssignments) {
                                Invoke-PnPCommandWithThrottling -Command {
                                    Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings -ErrorAction SilentlyContinue
                                } -OperationDescription "Load role assignment properties" | Out-Null

                                if (($ra.Member.LoginName -eq $userWithPerms.LoginName -or
                                        $ra.Member.Email -eq $userWithPerms.Email -or
                                        $ra.Member.UserPrincipalName -eq $userWithPerms.UserPrincipalName) -and
                                    $ra.Member.PrincipalType -eq "User") {
                                    foreach ($roleDef in $ra.RoleDefinitionBindings) {
                                        if ($roleDef.Name -ne "Limited Access") {
                                            $permLevels += $roleDef.Name
                                            Write-DebugInfo "  Direct permission: $($roleDef.Name)" -ForegroundColor DarkCyan
                                        }
                                    }
                                }
                            }

                            if ($permLevels.Count -gt 0) {
                                $permList = ($permLevels | Select-Object -Unique) -join ", "
                                $accessType = "Direct Access: $permList"
                                $userFound = $true
                                Write-DebugInfo "Found $user with direct access ($permList) on '$($site.Url)'" -ForegroundColor Cyan
                            }
                            else {
                                # No direct web-level role assignment found for this principal.
                                # Get-PnPUser -WithRightsAssigned also returns users who only have
                                # effective rights via a group (SP group, M365 group, Entra ID group).
                                # Do NOT mark as "Direct Access" here — let Checks 4/5 find the
                                # correct access path via group membership.
                                Write-DebugInfo "User appears in WithRightsAssigned but has no direct web-level role assignment (access likely via a group)" -ForegroundColor DarkGray
                            }
                        }
                        catch {
                            # Could not read web role assignments — do not assume "Direct Access".
                            # Let subsequent group checks determine the actual access path.
                            Write-DebugInfo "Could not read web role assignments for direct-access check: $($_.Exception.Message)" -ForegroundColor DarkYellow
                        }
                    }
                    else {
                        Write-DebugInfo "User not in rights-assigned list" -ForegroundColor DarkGray
                    }
                }
                catch {
                    Write-DebugInfo "Could not check rights assigned: $($_.Exception.Message)" -ForegroundColor DarkGray
                }

                # ── Check 1.5: M365 Group-connected site ────────────────────────────────────
                if ($site.GroupId -and $site.GroupId -ne "00000000-0000-0000-0000-000000000000") {
                    Write-DebugInfo "Checking M365 Group membership (ID: $($site.GroupId))..." -ForegroundColor Yellow
                    try {
                        $grpMembers = Invoke-PnPCommandWithThrottling -Command {
                            Get-PnPMicrosoft365GroupMember -Identity $site.GroupId -ErrorAction SilentlyContinue
                        } -OperationDescription "Get M365 group members"
                        if ($grpMembers) {
                            $isMember = $grpMembers | Where-Object { $_.UserPrincipalName -eq $user -or $_.Mail -eq $user }
                            if ($isMember) {
                                $userFound = $true
                                $accessType += "; M365 Group Member: $($site.Title)"
                                $groupMemberships += "M365 Group: $($site.Title)"
                                Write-DebugInfo "✓ $user is member of connected M365 group" -ForegroundColor Green
                            }
                        }
                        $grpOwners = Invoke-PnPCommandWithThrottling -Command {
                            Get-PnPMicrosoft365GroupOwner -Identity $site.GroupId -ErrorAction SilentlyContinue
                        } -OperationDescription "Get M365 group owners"
                        if ($grpOwners) {
                            $isOwner = $grpOwners | Where-Object { $_.UserPrincipalName -eq $user -or $_.Mail -eq $user }
                            if ($isOwner) {
                                $userFound = $true
                                $accessType += "; M365 Group Owner: $($site.Title)"
                                $groupMemberships += "M365 Group Owner: $($site.Title)"
                                Write-DebugInfo "✓ $user is OWNER of connected M365 group" -ForegroundColor Magenta
                            }
                        }
                    }
                    catch { Write-DebugInfo "Could not check M365 group membership: $($_.Exception.Message)" -ForegroundColor DarkYellow }
                }

                # ── Check 2: Site Owner ─────────────────────────────────────────────────────
                if ($site.Owner -eq $user) {
                    $userFound = $true
                    $accessType += "; Site Owner"
                    Write-DebugInfo "Found $user as site owner on '$($site.Url)'" -ForegroundColor Cyan
                }

                # ── Check 3: Site Collection Administrator ──────────────────────────────────
                try {
                    Write-DebugInfo "Checking Site Collection Administrators..." -ForegroundColor DarkYellow
                    $scAdmins = Invoke-PnPCommandWithThrottling -Command {
                        Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
                    } -OperationDescription "Get site collection admins"

                    if ($scAdmins) {
                        $isAdmin = $false
                        foreach ($adm in $scAdmins) {
                            if ($adm.LoginName -like "c:0t.c|tenant|*" -or $adm.LoginName -like "c:0o.c|federateddirectoryclaimprovider|*") {
                                try {
                                    $resolved = Get-PnPUser -Identity $adm.LoginName -ErrorAction SilentlyContinue
                                    if ($resolved -and ($resolved.UserPrincipalName -eq $user -or $resolved.Email -eq $user)) {
                                        $isAdmin = $true; break
                                    }
                                }
                                catch {}
                            }
                            elseif ($adm.LoginName -eq $user -or $adm.Email -eq $user -or
                                $adm.UserPrincipalName -eq $user -or $adm.LoginName -like "*$user*" -or
                                $adm.UserPrincipalName -like "*$user*") {
                                $isAdmin = $true; break
                            }
                        }
                        if ($isAdmin) {
                            $userFound = $true
                            $accessType += "; Site Collection Admin"
                            Write-DebugInfo "✓ Found $user as Site Collection Admin on '$($site.Url)'" -ForegroundColor Cyan
                        }
                    }
                }
                catch {
                    Write-DebugInfo "Error checking site collection admins: $($_.Exception.Message)" -ForegroundColor Red
                }

                # ── Check 4: SharePoint Group membership ────────────────────────────────────
                $siteGroups = Invoke-PnPCommandWithThrottling -Command {
                    Get-PnPGroup -ErrorAction SilentlyContinue
                } -OperationDescription "Get SharePoint groups"

                foreach ($grp in $siteGroups) {
                    try {
                        if ($grp.Title -like "*Limited Access*") { continue }
                        $members = @(Invoke-PnPCommandWithThrottling -Command {
                                Get-PnPGroupMember -Identity $grp.Title -ErrorAction SilentlyContinue
                            } -OperationDescription "Get members of SP group $($grp.Title)")
                        if ($members | Where-Object {
                                $e = $null; try { $e = $_.Email } catch {}
                                $_.LoginName -eq $user -or $e -eq $user
                            }) {
                            $userFound = $true
                            $accessType += "; SharePoint Group: $($grp.Title)"
                            $groupMemberships += $grp.Title
                            Write-DebugInfo "Found $user in SP group '$($grp.Title)'" -ForegroundColor Cyan
                        }
                    }
                    catch {}
                }

                # ── Check 5: Entra ID group membership ─────────────────────────────────────
                try {
                    $siteUsers = Invoke-PnPCommandWithThrottling -Command {
                        Get-PnPUser -ErrorAction SilentlyContinue
                    } -OperationDescription "Get site users for Entra ID group check"

                    if ($siteUsers) {
                        $usersWithRights2 = Invoke-PnPCommandWithThrottling -Command {
                            Get-PnPUser -WithRightsAssigned -ErrorAction SilentlyContinue
                        } -OperationDescription "Get users with rights for verification"

                        foreach ($siteUser in $siteUsers) {
                            if ($siteUser.PrincipalType -eq "SecurityGroup") {
                                try {
                                    $grpLoginName = $siteUser.LoginName
                                    $grpTitle = $siteUser.Title
                                    $grpId = $null

                                    if ($grpLoginName -like "c:0t.c|tenant|*") {
                                        $grpId = $grpLoginName -replace "c:0t.c\|tenant\|", ""
                                        Write-DebugInfo "Found Entra ID group (tenant claim): '$grpTitle' (ID: $grpId)" -ForegroundColor Yellow
                                    }
                                    elseif ($grpLoginName -like "c:0o.c|federateddirectoryclaimprovider|*") {
                                        $grpId = $grpLoginName -replace "c:0o.c\|federateddirectoryclaimprovider\|", ""
                                        Write-DebugInfo "Found Entra ID group (federated claim): '$grpTitle' (ID: $grpId)" -ForegroundColor Yellow
                                    }

                                    if ($grpId) {
                                        $isMember = Test-EntraGroupMembership -UserPrincipalName $user -GroupId $grpId -GroupDisplayName $grpTitle
                                        if ($isMember) {
                                            try {
                                                $web2 = Invoke-PnPCommandWithThrottling -Command {
                                                    Get-PnPWeb -ErrorAction SilentlyContinue
                                                } -OperationDescription "Get web for group role assignments"
                                                $ras = Invoke-PnPCommandWithThrottling -Command {
                                                    Get-PnPProperty -ClientObject $web2 -Property RoleAssignments -ErrorAction SilentlyContinue
                                                } -OperationDescription "Get role assignments for group permissions"
                                                $grpPerms = @()
                                                foreach ($ra in $ras) {
                                                    Invoke-PnPCommandWithThrottling -Command {
                                                        Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings -ErrorAction SilentlyContinue
                                                    } -OperationDescription "Load role assignment props for group" | Out-Null
                                                    if ($ra.Member.LoginName -eq $grpLoginName) {
                                                        foreach ($rd in $ra.RoleDefinitionBindings) {
                                                            if ($rd.Name -ne "Limited Access") { $grpPerms += $rd.Name }
                                                        }
                                                    }
                                                }
                                                $userFound = $true
                                                if ($grpPerms.Count -gt 0) {
                                                    $permList = ($grpPerms | Select-Object -Unique) -join ", "
                                                    $accessType += "; Entra ID Group: $grpTitle ($permList)"
                                                    $groupMemberships += "Entra ID: $grpTitle ($permList)"
                                                    Write-DebugInfo "✓ $user in Entra group '$grpTitle' ($permList)" -ForegroundColor Green
                                                }
                                                else {
                                                    $accessType += "; Entra ID Group: $grpTitle"
                                                    $groupMemberships += "Entra ID: $grpTitle"
                                                    Write-DebugInfo "✓ $user in Entra group '$grpTitle'" -ForegroundColor Green
                                                }
                                            }
                                            catch {
                                                $userFound = $true
                                                $accessType += "; Entra ID Group: $grpTitle"
                                            }
                                        }
                                    }
                                }
                                catch { Write-DebugInfo "Could not check Entra ID group '$($siteUser.Title)'" -ForegroundColor DarkYellow }
                            }
                        }
                    }
                }
                catch { Write-DebugInfo "Error getting site users: $($_.Exception.Message)" -ForegroundColor Red }

                # ── Check 6: Everyone except external users (optional) ──────────────────────
                if ($checkEEEU) {
                    try {
                        Write-DebugInfo "Checking 'Everyone except external users'..." -ForegroundColor Yellow
                        $EEEU = '*spo-grid-all-users*'
                        $eeeuInSite = $false
                        try {
                            $suWithRights = Invoke-PnPCommandWithThrottling -Command {
                                Get-PnPUser -WithRightsAssigned -ErrorAction SilentlyContinue
                            } -OperationDescription "Get site users with rights for EEEU check"
                            if ($suWithRights | Where-Object { $_.LoginName -like $EEEU }) { $eeeuInSite = $true }
                        }
                        catch {}

                        if ($eeeuInSite) {
                            $perms = Invoke-PnPCommandWithThrottling -Command {
                                Get-PnPProperty -ClientObject (Get-PnPWeb) -Property RoleAssignments -ErrorAction SilentlyContinue
                            } -OperationDescription "Get web role assignments for EEEU check"

                            if ($perms) {
                                foreach ($ra in $perms) {
                                    try {
                                        Invoke-PnPCommandWithThrottling -Command {
                                            Get-PnPProperty -ClientObject $ra -Property Member -ErrorAction SilentlyContinue | Out-Null
                                            Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings -ErrorAction SilentlyContinue | Out-Null
                                        } -OperationDescription "Get EEEU role assignment props" | Out-Null
                                        if ($ra.Member.LoginName -like $EEEU -and $ra.RoleDefinitionBindings.Name -ne 'Limited Access') {
                                            if ($user -like "*@$t.onmicrosoft.com" -or $user -like "*@*.onmicrosoft.com") {
                                                $userFound = $true
                                                $accessType += "; Everyone except external users"
                                                Write-DebugInfo "✓ $user has access via EEEU" -ForegroundColor Green
                                            }
                                            break
                                        }
                                    }
                                    catch {}
                                }
                            }
                        }
                    }
                    catch { Write-Host "Error checking EEEU: $($_.Exception.Message)" -ForegroundColor Red }
                }

                # ── Check 7: User Information List fallback ─────────────────────────────────
                # Get-PnPUser (no flags) returns every principal ever recorded in the UIL, including
                # residual entries left after access is removed and phantom entries auto-created when
                # a user visits via EEEU.  We can't reliably distinguish these two cases from the UIL
                # entry alone, so we ALWAYS report the finding but qualify the AccessType with
                # context so the admin can decide whether to act on it:
                #
                #   "UIL Entry (residual)"
                #       – user is in UIL, has no current rights, EEEU is NOT active.
                #         Almost certainly a genuine leftover entry that should be cleaned up.
                #
                #   "UIL Entry (residual – EEEU also active, verify not phantom)"
                #       – user is in UIL, has no current rights, but EEEU IS active on the site.
                #         Could be a residual entry OR a phantom created by an EEEU visit.
                #         Admin should verify before deciding whether to remove.
                if (-not $userFound) {
                    Write-DebugInfo "No explicit permissions found. Checking User Information List (UIL)..." -ForegroundColor Yellow
                    try {
                        $allUsersUIL = Invoke-PnPCommandWithThrottling -Command {
                            Get-PnPUser -ErrorAction SilentlyContinue
                        } -OperationDescription "Get all site users (UIL check)"

                        if ($allUsersUIL) {
                            $inList = $allUsersUIL | Where-Object {
                                $_.LoginName -eq $user -or $_.Email -eq $user -or
                                $_.UserPrincipalName -eq $user -or $_.LoginName -like "*$user*"
                            }

                            if ($inList) {
                                Write-DebugInfo "  User found in UIL - checking EEEU status to qualify the finding..." -ForegroundColor Yellow

                                # Check whether EEEU has active permissions on this site
                                $eeeuActive = $false
                                try {
                                    $withRightsForEEEU = Invoke-PnPCommandWithThrottling -Command {
                                        Get-PnPUser -WithRightsAssigned -ErrorAction SilentlyContinue
                                    } -OperationDescription "Check EEEU presence for UIL qualification"
                                    $eeeuActive = ($null -ne ($withRightsForEEEU | Where-Object { $_.LoginName -like '*spo-grid-all-users*' }))
                                }
                                catch { Write-DebugInfo "  Could not check EEEU presence: $($_.Exception.Message)" -ForegroundColor DarkYellow }

                                $userFound = $true
                                if ($eeeuActive) {
                                    $accessType = "UIL Entry (residual - EEEU also active, verify not phantom)"
                                    Write-StatusMessage "⚠️  Found $user in UIL on '$($site.Url)' - EEEU is active; entry may be residual or phantom. Verify before removing." -ForegroundColor Yellow
                                    Write-LogEntry -LogName $log -LogEntryText "Found $user in UIL on '$($site.Url)' - EEEU active; entry may be phantom or residual (no explicit rights assigned)"
                                }
                                else {
                                    $accessType = "UIL Entry (residual)"
                                    Write-StatusMessage "✓ Found $user in UIL on '$($site.Url)' - residual entry, no active permissions, EEEU not active" -ForegroundColor Magenta
                                    Write-LogEntry -LogName $log -LogEntryText "Found $user in UIL on '$($site.Url)' - residual entry only (no rights assigned, EEEU not active)"
                                }
                            }
                        }
                    }
                    catch { Write-DebugInfo "Could not check UIL: $($_.Exception.Message)" -ForegroundColor DarkYellow }
                }

                # ── Write result ────────────────────────────────────────────────────────────
                if ($userFound) {
                    $cleanAccess = ($accessType.TrimStart('; ').Trim() -split '; ') |
                    Where-Object { $_ -notlike "*Limited Access*" }
                    $finalAccess = $cleanAccess -join '; '

                    $row = [PSCustomObject]@{
                        SiteName   = $site.Title
                        URL        = $site.Url
                        User       = $user
                        Owner      = $site.Owner
                        AccessType = $finalAccess
                    }
                    $siteOutput += $row
                    Write-LogEntry -LogName $log -LogEntryText "Found $user on '$($site.Url)' - $finalAccess"
                }
                else {
                    Write-DebugInfo "$user NOT FOUND on '$($site.Url)'" -ForegroundColor Magenta
                    Write-LogEntry -LogName $log -LogEntryText "$user NOT FOUND on '$($site.Url)'"
                }
            }
            catch {
                Write-DebugInfo "$user NOT FOUND on '$($site.Url)' (error)" -ForegroundColor Magenta
                Write-LogEntry -LogName $log -LogEntryText "$user NOT FOUND on '$($site.Url)' - $($_.Exception.Message)"
            }
            Write-Host ""
            Write-LogEntry -LogName $log -LogEntryText ""
        }

        # Write site results to CSV
        if ($siteOutput.Count -gt 0) {
            if ($script:firstWrite) {
                $siteOutput | Export-Csv $outputFile -NoTypeInformation
                $script:firstWrite = $false
                Write-DebugInfo "Exported $($siteOutput.Count) user(s) for '$($site.Title)' (with headers)" -ForegroundColor Green
            }
            else {
                $siteOutput | Export-Csv $outputFile -NoTypeInformation -Append
                Write-DebugInfo "Exported $($siteOutput.Count) user(s) for '$($site.Title)' (appended)" -ForegroundColor Green
            }
            Write-LogEntry -LogName $log -LogEntryText "Exported $($siteOutput.Count) user(s) for '$($site.Title)' to CSV"

            # ── Inline removal (Both mode) ──────────────────────────────────────────────
            if ($RemoveInline) {
                $foundUsers = $siteOutput | Select-Object -ExpandProperty User | Sort-Object -Unique
                Write-StatusMessage "`nBOTH mode: Removing $($foundUsers.Count) user(s) found on '$($site.Title)'..." -ForegroundColor Magenta
                Write-LogEntry -LogName $log -LogEntryText "BOTH mode: inline removal of [$($foundUsers -join ', ')] from $($site.Url)"
                Invoke-ProcessSingleSite -SiteUrl $site.Url -Users $foundUsers
            }

            $siteOutput = @()
        }

        Write-StatusMessage "Completed processing site: $($site.Title)" -ForegroundColor Yellow
        Write-DebugInfo "----------------------------------------" -ForegroundColor DarkGray
        Write-LogEntry -LogName $log -LogEntryText "Completed processing site: $($site.Title)"
        Write-LogEntry -LogName $log -LogEntryText "----------------------------------------"
    }

    Write-Host ""
    Write-StatusMessage "All sites processed. Report saved to: $outputFile" -ForegroundColor Green
    Write-StatusMessage "Log file: $log" -ForegroundColor Green
    return $outputFile
}

#endregion ────────────────────────────────────────────────────────────────────────────────────────

#region ── REMOVE Functions ───────────────────────────────────────────────────────────────────────

function Remove-UserFromSiteGroups {
    param([array]$Users)
    Write-Host "`nRemoving users from site groups..." -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting removal from site groups"
    try {
        $siteGroups = Invoke-WithThrottleHandling -ScriptBlock { Get-PnPGroup } -Operation "Get site groups"
        Write-Host "Found $($siteGroups.Count) groups" -ForegroundColor Green
        foreach ($grp in $siteGroups) {
            # SharingLinks groups are handled exclusively by Remove-UserFromSharingLinks
            if ($grp.Title -match '^SharingLinks\.') { continue }
            Write-Host "  Group: $($grp.Title)" -ForegroundColor Yellow
            try {
                $members = @(Invoke-WithThrottleHandling -ScriptBlock {
                        Get-PnPGroupMember -Identity $grp.Id
                    } -Operation "Get members for group $($grp.Title)")

                foreach ($user in $Users) {
                    $match = $members | Where-Object {
                        $e = $null; try { $e = $_.Email } catch {}
                        $_.LoginName -eq $user -or $_.LoginName -like "*$user*" -or $e -eq $user
                    }
                    if ($match) {
                        try {
                            Invoke-WithThrottleHandling -ScriptBlock {
                                try { Remove-PnPGroupMember -Identity $grp.Id -LoginName $match.LoginName -Force }
                                catch { Remove-PnPGroupMember -Identity $grp.Id -LoginName $match.LoginName }
                            } -Operation "Remove $user from group $($grp.Title)"
                            Write-Host "    Removed $user from group: $($grp.Title)" -ForegroundColor Green
                            Write-LogEntry -LogName $log -LogEntryText "Removed $user from group: $($grp.Title)"
                        }
                        catch {
                            Write-Host "    Error removing $user from group $($grp.Title): $_" -ForegroundColor Red
                            Write-LogEntry -LogName $log -LogEntryText "ERROR removing $user from group $($grp.Title): $_"
                        }
                    }
                }
            }
            catch {
                Write-Host "    Error processing group $($grp.Title): $_" -ForegroundColor Red
                Write-LogEntry -LogName $log -LogEntryText "ERROR processing group $($grp.Title): $_"
            }
        }
    }
    catch {
        Write-Host "Error processing site groups: $_" -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText "ERROR processing site groups: $_"
    }
}

function Remove-UserFromFilePermissions {
    param([array]$Users)
    Write-Host "`nRemoving users from file/item permissions..." -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting removal from file/item permissions"
    try {
        $lists = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPList | Where-Object { $_.Hidden -eq $false -and ($_.BaseType -eq "DocumentLibrary" -or $_.BaseType -eq "GenericList") }
        } -Operation "Get document libraries and lists"

        Write-Host "Found $($lists.Count) libraries/lists" -ForegroundColor Green
        foreach ($list in $lists) {
            $listType = if ($list.BaseType -eq "DocumentLibrary") { "library" } else { "list" }
            Write-Host "  Processing $listType : $($list.Title)" -ForegroundColor Yellow
            try {
                $items = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPListItem -List $list.Id -PageSize 1000
                } -Operation "Get items from $listType $($list.Title)"

                foreach ($item in $items) {
                    try {
                        # Load the HasUniqueRoleAssignments property without letting Get-PnPProperty
                        # bleed extra output into the pipeline (which would make $hasUnique always truthy).
                        Invoke-WithThrottleHandling -ScriptBlock {
                            Get-PnPProperty -ClientObject $item -Property "HasUniqueRoleAssignments" | Out-Null
                        } -Operation "Check unique perms for item $($item.Id)"
                        $hasUnique = $item.HasUniqueRoleAssignments

                        if (-not $hasUnique) {
                            # Item inherits permissions from its parent – nothing to remove here.
                            # The user will be removed at the appropriate level (site group / site permissions).
                            Write-DebugInfo "    Skipping item $($item.Id) in '$($list.Title)' - permissions are inherited" -ForegroundColor DarkGray
                            Write-LogEntry -LogName $log -LogEntryText "Skipped item $($item.Id) in '$($list.Title)' - inherits permissions (no unique perms to remove)"
                            continue
                        }

                        if ($hasUnique) {
                            try {
                                Invoke-WithThrottleHandling -ScriptBlock {
                                    $item.Context.Load($item.RoleAssignments)
                                    $item.Context.ExecuteQuery()
                                    foreach ($ra in $item.RoleAssignments) {
                                        $item.Context.Load($ra.Member)
                                        $item.Context.Load($ra.RoleDefinitionBindings)
                                    }
                                    $item.Context.ExecuteQuery()
                                } -Operation "Load role assignments for item $($item.Id)"

                                foreach ($ra in $item.RoleAssignments) {
                                    try {
                                        $member = $ra.Member
                                        foreach ($user in $Users) {
                                            # Email property only exists on User principals—SP groups and system accounts throw if accessed directly.
                                            $memberEmail = try { $member.Email } catch { $null }
                                            if ($member.LoginName -eq $user -or $member.LoginName -like "*$user*" -or $memberEmail -eq $user) {
                                                try {
                                                    Invoke-WithThrottleHandling -ScriptBlock {
                                                        $removedRoles = @()
                                                        foreach ($rd in $ra.RoleDefinitionBindings) {
                                                            try {
                                                                if ($rd.Name -ne "Limited Access") {
                                                                    Set-PnPListItemPermission -List $list.Id -Identity $item.Id -User $member.LoginName -RemoveRole $rd.Name
                                                                    $removedRoles += $rd.Name
                                                                }
                                                            }
                                                            catch {
                                                                if ($_.Exception.Message -notlike "*Can not find the principal*" -and $_.Exception.Message -notlike "*does not exist*") {
                                                                    Write-LogEntry -LogName $log -LogEntryText "WARN removing role '$($rd.Name)' for $user on $($item['FileLeafRef']): $_"
                                                                }
                                                            }
                                                        }
                                                        if ($removedRoles.Count -gt 0) {
                                                            Write-LogEntry -LogName $log -LogEntryText "Removed roles [$($removedRoles -join ', ')] for $user on $($item['FileLeafRef'])"
                                                        }
                                                    } -Operation "Remove $user from item $($item.Id)"

                                                    $displayName = if ($list.BaseType -eq "DocumentLibrary") { $item["FileLeafRef"] } else { "Item $($item.Id) ($($item['Title']))" }
                                                    Write-Host "    Removed $user from ${listType}: $displayName" -ForegroundColor Green
                                                    Write-LogEntry -LogName $log -LogEntryText "Removed $user from $displayName in $($list.Title)"
                                                }
                                                catch {
                                                    if ($_.Exception.Message -notlike "*Can not find the principal*" -and $_.Exception.Message -notlike "*does not exist*") {
                                                        Write-Host "    Error removing $user from $($item['FileLeafRef']): $_" -ForegroundColor Red
                                                        Write-LogEntry -LogName $log -LogEntryText "ERROR removing $user from $($item['FileLeafRef']): $_"
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { Write-LogEntry -LogName $log -LogEntryText "Error processing role assignment member: $_" }
                                }
                            }
                            catch { Write-LogEntry -LogName $log -LogEntryText "Error loading role assignments for item $($item.Id): $_" }

                            # ── Flexible sharing link invitation cleanup ─────────────────────────
                            # GetSharingInformation is the authoritative source for Flexible link
                            # invitees. It catches orphaned invitations even when the backing SP
                            # group has already been deleted by a prior run or manual action.
                            # ShareLink with inviteesToRemove removes BOTH the SP group member AND
                            # the invitation metadata that the UI reads.
                            try {
                                $siUrl = "/_api/web/Lists('$($list.Id)')/GetItemById($($item.Id))/GetSharingInformation?`$expand=permissionsInformation,sharingLinkTemplates"
                                $si = Invoke-PnPSPRestMethod -Method Post -Url $siUrl `
                                    -Content @{ request = @{ maxPrincipalsToReturn = 100; maxLinkMembersToReturn = 100 } } `
                                    -ErrorAction SilentlyContinue

                                $flexLinks = if ($si) {
                                    @($si.permissionsInformation.links) | Where-Object {
                                        $_.linkDetails -and [int]$_.linkDetails.LinkKind -eq 6 -and $_.linkDetails.IsActive -eq $true
                                    }
                                }
                                else { @() }
                                $templates = if ($si) { @($si.sharingLinkTemplates.templates) } else { @() }

                                foreach ($flexLink in $flexLinks) {
                                    $shareId = $flexLink.linkDetails.ShareId
                                    $invites = @($flexLink.linkDetails.Invitations)
                                    $members = @($flexLink.linkMembers)

                                    # Look up the actual stored role once per link – never assume
                                    $linkRole = $null
                                    $matchTpl = $templates | Where-Object { $_.linkDetails -and $_.linkDetails.ShareId -ieq $shareId }
                                    if ($matchTpl) { $linkRole = [int]$matchTpl.role }
                                    if ($null -eq $linkRole) {
                                        Write-LogEntry -LogName $log -LogEntryText "Could not determine role for Flexible link shareId '$shareId' on item $($item.Id) – skipping"
                                        continue
                                    }

                                    foreach ($user in $Users) {
                                        # Check invitations list first; fall back to linkMembers
                                        $inviteeInfo = $null
                                        $fromInvite = $invites | Where-Object { $_.invitee -and ($_.invitee.email -ieq $user -or $_.invitee.userPrincipalName -ieq $user) }
                                        $fromMember = $members | Where-Object { $_.email -ieq $user -or $_.userPrincipalName -ieq $user }
                                        if ($fromInvite) { $inviteeInfo = $fromInvite.invitee }
                                        elseif ($fromMember) { $inviteeInfo = $fromMember }
                                        if (-not $inviteeInfo) { continue }

                                        $inviteeId = 0; try { $inviteeId = [int]$inviteeInfo.id } catch {}
                                        $inviteeName = ""; try { $inviteeName = $inviteeInfo.name } catch {}
                                        $inviteeLogin = ""; try { $inviteeLogin = $inviteeInfo.loginName } catch {}
                                        $inviteeEmail = ""; try { $inviteeEmail = $inviteeInfo.email } catch {}

                                        $shareLinkBody = @{
                                            request = @{
                                                createLink = $true
                                                settings   = @{
                                                    linkKind                = 6
                                                    expiration              = $null
                                                    role                    = $linkRole
                                                    restrictShareMembership = $true
                                                    shareId                 = $shareId
                                                    scope                   = 2
                                                    nav                     = ""
                                                    inviteesToRemove        = @(
                                                        @{
                                                            id            = $inviteeId
                                                            loginName     = $inviteeLogin
                                                            name          = $inviteeName
                                                            isExternal    = $false
                                                            principalType = 1
                                                            email         = $inviteeEmail
                                                        }
                                                    )
                                                }
                                                emailData  = @{ body = "" }
                                            }
                                        }

                                        $slUrl = "/_api/web/Lists(@a1)/GetItemById(@a2)/ShareLink?@a1='$($list.Id)'&@a2='$($item.Id)'"
                                        Invoke-PnPSPRestMethod -Method Post -Url $slUrl -Content $shareLinkBody -ErrorAction Stop | Out-Null

                                        $displayName = if ($list.BaseType -eq "DocumentLibrary") { $item["FileLeafRef"] } else { "Item $($item.Id)" }
                                        Write-Host "    Removed $user from Flexible sharing link (shareId: $shareId) on '$displayName'" -ForegroundColor Green
                                        Write-LogEntry -LogName $log -LogEntryText "Removed $user from Flexible sharing link via ShareLink API (shareId: $shareId, role: $linkRole) on '$displayName'"
                                    }
                                }
                            }
                            catch {
                                Write-LogEntry -LogName $log -LogEntryText "Flexible link invitation check failed for item $($item.Id) in '$($list.Title)': $_"
                            }
                        }
                    }
                    catch { Write-LogEntry -LogName $log -LogEntryText "Error processing item $($item.Id): $_" }
                }
            }
            catch {
                Write-Host "  Error processing $listType $($list.Title): $_" -ForegroundColor Red
                Write-LogEntry -LogName $log -LogEntryText "ERROR processing $listType $($list.Title): $_"
            }
        }
    }
    catch {
        Write-Host "Error processing file permissions: $_" -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText "ERROR processing file permissions: $_"
    }
}

function Remove-UserFromSharingLinks {
    param([array]$Users)
    Write-Host "`nRemoving users from sharing link groups..." -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting removal from sharing link groups"

    try {
        $allGroups = @(Invoke-WithThrottleHandling -ScriptBlock { Get-PnPGroup } -Operation "Get all groups")
        $sharingGroups = @($allGroups | Where-Object { $_.Title -match '^SharingLinks\.' })
        Write-Host "Found $($sharingGroups.Count) sharing link group(s)" -ForegroundColor Green

        foreach ($grp in $sharingGroups) {
            Write-Host "  Group: $($grp.Title)" -ForegroundColor Yellow
            try {
                $members = @(Invoke-WithThrottleHandling -ScriptBlock {
                        Get-PnPGroupMember -Identity $grp.Id -ErrorAction SilentlyContinue
                    } -Operation "Get members for $($grp.Title)")

                foreach ($user in $Users) {
                    # Locate this user in the group members
                    $match = $null
                    foreach ($m in $members) {
                        $mLogin = $null; try { $mLogin = $m.LoginName } catch {}
                        $mEmail = $null; try { $mEmail = $m.Email } catch {}
                        if ($mLogin -eq $user -or $mLogin -like "*$user*" -or $mEmail -eq $user) {
                            $match = $m; break
                        }
                    }
                    if (-not $match) { continue }

                    # Determine whether this is a Flexible (specific-people) link
                    if ($grp.Title -match '^SharingLinks\.([0-9a-fA-F\-]{36})\.Flexible\.([0-9a-fA-F\-]{36})$') {
                        $itemUniqueId = $Matches[1]   # First GUID  = the file/item UniqueId
                        $shareId = $Matches[2]   # Second GUID = the sharing link ID (ShareId)

                        # Use the same REST API the SharePoint UI calls:
                        #   POST /_api/web/Lists(@a1)/GetItemById(@a2)/ShareLink
                        # with inviteesToRemove – this removes the user from BOTH the SP group
                        # membership AND the invitation metadata that the UI displays.
                        $removed = $false
                        try {
                            # Step 1: Resolve the list GUID and list-item integer ID from the file UniqueId.
                            $fileDetails = Invoke-PnPSPRestMethod `
                                -Url "/_api/web/GetFileById('$itemUniqueId')?`$select=ListId,ListItemAllFields/Id&`$expand=ListItemAllFields" `
                                -ErrorAction Stop

                            $listId = $fileDetails.ListId
                            $spItemId = $fileDetails.ListItemAllFields.Id
                            Write-LogEntry -LogName $log -LogEntryText "Resolved list/item for group '$($grp.Title)': listId=$listId  itemId=$spItemId"

                            # Step 2: Look up the actual role stored on this sharing link via
                            # GetSharingInformation – never hard-code the role assumption.
                            $linkRole = $null
                            try {
                                $sharingInfoUrl = "/_api/web/Lists('$listId')/GetItemById($spItemId)/GetSharingInformation?`$expand=sharingLinkTemplates"
                                $sharingInfo = Invoke-PnPSPRestMethod -Method Post -Url $sharingInfoUrl `
                                    -Content @{ request = @{ maxPrincipalsToReturn = 1; maxLinkMembersToReturn = 1 } } `
                                    -ErrorAction Stop
                                # Non-verbose JSON: arrays are plain arrays, no .results wrapper
                                $templates = @($sharingInfo.sharingLinkTemplates.templates)
                                $matchingTpl = $templates | Where-Object {
                                    $_.linkDetails -and $_.linkDetails.ShareId -ieq $shareId
                                }
                                if ($matchingTpl) {
                                    $linkRole = [int]$matchingTpl.role
                                    Write-LogEntry -LogName $log -LogEntryText "Link role for shareId $shareId resolved to: $linkRole"
                                }
                            }
                            catch {
                                Write-LogEntry -LogName $log -LogEntryText "Could not look up link role for shareId $shareId (will abort rather than assume): $_"
                            }

                            if ($null -eq $linkRole) {
                                throw "Could not determine the role for shareId '$shareId' - aborting to avoid incorrect modification"
                            }

                            # Step 3: Build the invitee payload from the matched group member object.
                            $inviteeId = 0; try { $inviteeId = [int]$match.Id } catch {}
                            $inviteeName = ""; try { $inviteeName = $match.Title } catch {}
                            $inviteeLogin = ""; try { $inviteeLogin = $match.LoginName } catch {}
                            $inviteeEmail = ""; try { $inviteeEmail = $match.Email } catch {}

                            # Step 4: POST ShareLink with inviteesToRemove – mirrors exactly what the UI sends.
                            # inviteesToRemove must be a plain JSON array (not OData verbose {"results":[...]})
                            # because Invoke-PnPSPRestMethod uses non-verbose application/json serialization.
                            $shareLinkBody = @{
                                request = @{
                                    createLink = $true
                                    settings   = @{
                                        linkKind                = 6
                                        expiration              = $null
                                        role                    = $linkRole
                                        restrictShareMembership = $true
                                        shareId                 = $shareId
                                        scope                   = 2
                                        nav                     = ""
                                        inviteesToRemove        = @(
                                            @{
                                                id            = $inviteeId
                                                loginName     = $inviteeLogin
                                                name          = $inviteeName
                                                isExternal    = $false
                                                principalType = 1
                                                email         = $inviteeEmail
                                            }
                                        )
                                    }
                                    emailData  = @{ body = "" }
                                }
                            }

                            $shareLinkUrl = "/_api/web/Lists(@a1)/GetItemById(@a2)/ShareLink?@a1='$listId'&@a2='$spItemId'"
                            Invoke-PnPSPRestMethod -Method Post -Url $shareLinkUrl -Content $shareLinkBody -ErrorAction Stop | Out-Null

                            Write-Host "    Removed $user from Flexible sharing link (shareId: $shareId, role: $linkRole)" -ForegroundColor Green
                            Write-LogEntry -LogName $log -LogEntryText "Removed $user from Flexible sharing link via ShareLink API (shareId: $shareId, role: $linkRole, listId: $listId, itemId: $spItemId)"
                            $removed = $true
                        }
                        catch {
                            Write-Host "    ShareLink API failed - falling back to SP group removal: $_" -ForegroundColor Yellow
                            Write-LogEntry -LogName $log -LogEntryText "ShareLink API failed for $user (shareId: $shareId): $_ - falling back to SP group removal"
                        }

                        if (-not $removed) {
                            # Fallback: remove from the SP group only (invitee metadata may persist in UI)
                            try {
                                $loginToRemove = $null; try { $loginToRemove = $match.LoginName } catch {}
                                Remove-PnPGroupMember -Identity $grp.Id -LoginName $loginToRemove -ErrorAction Stop
                                Write-LogEntry -LogName $log -LogEntryText "Fallback: removed $user from SP group '$($grp.Title)' (UI entry may persist)"
                            }
                            catch {
                                Write-LogEntry -LogName $log -LogEntryText "Fallback SP group removal also failed for '$($grp.Title)': $_"
                            }
                        }

                    }
                    else {
                        # Non-Flexible sharing group (OrganizationEdit, OrganizationView, etc.)
                        $loginToRemove = $null; try { $loginToRemove = $match.LoginName } catch {}
                        try {
                            Invoke-WithThrottleHandling -ScriptBlock {
                                Remove-PnPGroupMember -Identity $grp.Id -LoginName $loginToRemove -ErrorAction Stop
                            } -Operation "Remove $user from sharing group $($grp.Title)"
                            Write-Host "    Removed $user from sharing group: $($grp.Title)" -ForegroundColor Green
                            Write-LogEntry -LogName $log -LogEntryText "Removed $user from sharing group: $($grp.Title)"
                        }
                        catch {
                            Write-Host "    Error removing $user from $($grp.Title): $_" -ForegroundColor Red
                            Write-LogEntry -LogName $log -LogEntryText "ERROR removing $user from sharing group '$($grp.Title)': $_"
                        }
                    }
                }
            }
            catch {
                Write-Host "  Error processing group $($grp.Title): $_" -ForegroundColor Red
                Write-LogEntry -LogName $log -LogEntryText "ERROR processing sharing group '$($grp.Title)': $_"
            }
        }
    }
    catch {
        Write-Host "Error processing sharing links: $_" -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText "ERROR processing sharing links: $_"
    }
}

function Invoke-ProcessSingleSite {
    param([string]$SiteUrl, [array]$Users)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Processing Site: $SiteUrl" -ForegroundColor Cyan
    Write-Host "Users to remove: $($Users -join ', ')" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting REMOVE processing for site: $SiteUrl  Users: $($Users -join ', ')"

    try {
        Connect-ToSite -SiteUrl $SiteUrl

        Remove-UserFromSiteGroups      -Users $Users
        Remove-UserFromFilePermissions -Users $Users
        Remove-UserFromSharingLinks    -Users $Users

        if ($RemoveFromUIL) {
            Write-Host "`nRemoving users from User Information List..." -ForegroundColor Cyan
            Write-LogEntry -LogName $log -LogEntryText "Starting UIL removal"
            foreach ($user in $Users) {
                try {
                    Write-Host "  Looking up $user..." -ForegroundColor Yellow
                    $pnpUser = Get-PnPUser | Where-Object { $_.Email -eq $user }
                    if ($pnpUser) {
                        $loginName = $pnpUser.LoginName
                        Write-Host "  Removing $loginName from UIL..." -ForegroundColor Yellow
                        Remove-PnPUser -Identity $loginName -Force:$true
                        Write-Host "  Successfully removed $user from UIL" -ForegroundColor Green
                        Write-LogEntry -LogName $log -LogEntryText "Removed $user (LoginName: $loginName) from UIL"
                    }
                    else {
                        Write-Host "  $user not found in site UIL" -ForegroundColor Yellow
                        Write-LogEntry -LogName $log -LogEntryText "User $user not found in UIL"
                    }
                }
                catch {
                    Write-Host "  Failed to remove $user from UIL: $_" -ForegroundColor Red
                    Write-LogEntry -LogName $log -LogEntryText "FAILED to remove $user from UIL: $_"
                }
            }
        }

        Write-Host "`nCompleted processing for site: $SiteUrl" -ForegroundColor Green
        Write-LogEntry -LogName $log -LogEntryText "Completed REMOVE for site: $SiteUrl"
    }
    catch {
        Write-Host "`nError processing site $SiteUrl : $_" -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText "ERROR processing site $SiteUrl : $_"
        throw
    }
    finally {
        try { Disconnect-PnPOnline } catch {}
    }
}

function Invoke-RemoveMode {
    param([string]$CsvPath)

    Write-StatusMessage "`n========================================" -ForegroundColor Magenta
    Write-StatusMessage "  REMOVE MODE  -  Using CSV: $CsvPath" -ForegroundColor Magenta
    Write-StatusMessage "========================================`n" -ForegroundColor Magenta
    Write-LogEntry -LogName $log -LogEntryText "Starting REMOVE mode. Input CSV: $CsvPath"

    if (-not (Test-Path $CsvPath)) {
        Write-Host "ERROR: Input CSV not found: $CsvPath" -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText "ERROR: Input CSV not found: $CsvPath"
        exit 1
    }

    $csvData = Import-Csv -Path $CsvPath
    if (-not $csvData -or $csvData.Count -eq 0) {
        Write-Host "ERROR: Input CSV is empty or invalid: $CsvPath" -ForegroundColor Red
        exit 1
    }

    # Validate required columns
    $requiredCols = @('URL', 'User')
    $csvCols = ($csvData | Get-Member -MemberType NoteProperty).Name
    foreach ($col in $requiredCols) {
        if ($col -notin $csvCols) {
            Write-Host "ERROR: CSV is missing required column '$col'. Expected columns: SiteName, URL, User, Owner, AccessType" -ForegroundColor Red
            exit 1
        }
    }

    # Group by site URL so we process each site once with ALL its affected users
    $bySite = $csvData | Group-Object -Property URL
    Write-StatusMessage "Found $($csvData.Count) row(s) across $($bySite.Count) unique site(s)" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "CSV contains $($csvData.Count) rows across $($bySite.Count) unique sites"

    $siteIndex = 0
    $successCount = 0
    $errorCount = 0

    foreach ($siteGroup in $bySite) {
        $siteIndex++
        $siteUrl = $siteGroup.Name
        $usersOnSite = $siteGroup.Group | Select-Object -ExpandProperty User | Sort-Object -Unique
        $siteName = ($siteGroup.Group | Select-Object -First 1).SiteName

        Write-Host "`n`n========================================" -ForegroundColor Magenta
        Write-Host "Site $siteIndex of $($bySite.Count): $siteName" -ForegroundColor Magenta
        Write-Host "========================================" -ForegroundColor Magenta

        try {
            Invoke-ProcessSingleSite -SiteUrl $siteUrl -Users $usersOnSite
            $successCount++
        }
        catch {
            Write-Host "Failed to process site: $siteUrl" -ForegroundColor Red
            Write-LogEntry -LogName $log -LogEntryText "Failed to process site: $siteUrl - $_"
            $errorCount++
        }
    }

    Write-Host "`n`n========================================" -ForegroundColor Magenta
    Write-Host "REMOVE MODE COMPLETE" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Sites processed : $($bySite.Count)" -ForegroundColor White
    Write-Host "Successful      : $successCount"    -ForegroundColor Green
    Write-Host "Failed          : $errorCount"      -ForegroundColor $(if ($errorCount -gt 0) { 'Red' } else { 'Green' })
    Write-Host "Log file        : $log"             -ForegroundColor White
    Write-LogEntry -LogName $log -LogEntryText "REMOVE MODE COMPLETE - Sites: $($bySite.Count), Success: $successCount, Failed: $errorCount"
}

#endregion ────────────────────────────────────────────────────────────────────────────────────────

#region ── Main Execution ─────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "  SharePoint Online Permissions Manager" -ForegroundColor Cyan
Write-Host "  Mode   : $Mode" -ForegroundColor Cyan
Write-Host "  Log    : $log" -ForegroundColor Cyan
if ($Mode -in @("Report", "Both")) {
    Write-Host "  Output : $outputFile" -ForegroundColor Cyan
}
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host ""

Write-LogEntry -LogName $log -LogEntryText "============================================"
Write-LogEntry -LogName $log -LogEntryText "SPO Permissions Manager started. Mode: $Mode"
Write-LogEntry -LogName $log -LogEntryText "============================================"

# ── Connect to tenant admin (required for Report and for site enumeration) ─────────────────────
if ($Mode -in @("Report", "Both")) {
    try {
        Write-StatusMessage "Connecting to SharePoint Online tenant admin..." -ForegroundColor Green
        Connect-PnPOnline -Url "https://$t-admin.sharepoint.com" -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        Write-StatusMessage "Connected to SharePoint Online" -ForegroundColor Green
        Write-LogEntry -LogName $log -LogEntryText "Connected to tenant admin: https://$t-admin.sharepoint.com"

        if ($enableThrottlingProtection) {
            Write-DebugInfo "Throttling protection ENABLED (delays: ${baseDelayBetweenSites}s/site, ${baseDelayBetweenUsers}s/user; max retries: $maxRetryAttempts)" -ForegroundColor Green
        }
        if ($debug) {
            Write-DebugInfo "DEBUG MODE: ENABLED" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "FATAL: Failed to connect to SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText "FATAL: Connection failed: $($_.Exception.Message)"
        exit 1
    }

    # Load users list (required for Report mode)
    if (-not (Test-Path $UsersFilePath)) {
        Write-Host "FATAL: Users file not found: $UsersFilePath" -ForegroundColor Red
        exit 1
    }
    $users = @(Get-Content $UsersFilePath | Where-Object { $_ -and $_.Trim() -ne "" })
    if ($users.Count -eq 0) {
        Write-Host "FATAL: No users found in $UsersFilePath" -ForegroundColor Red
        exit 1
    }
    Write-StatusMessage "Loaded $($users.Count) user(s) from $UsersFilePath" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Loaded $($users.Count) users from $UsersFilePath"
}

# ── Run REPORT ─────────────────────────────────────────────────────────────────────────────────
if ($Mode -in @("Report", "Both")) {
    # Both mode removes users inline per-site as each site's scan completes (no CSV round-trip).
    $generatedCsv = Invoke-ReportMode -Users $users -RemoveInline:($Mode -eq "Both")
}

# ── Run REMOVE ─────────────────────────────────────────────────────────────────────────────────
if ($Mode -in @("Remove", "Both")) {
    if ($Mode -eq "Both") {
        Write-StatusMessage "`nBOTH mode: User removals were performed inline during the Report scan. No separate Remove pass needed." -ForegroundColor Magenta
        Write-LogEntry -LogName $log -LogEntryText "BOTH mode complete - inline removal performed per site during Report scan. Separate Remove pass skipped."
    }
    else {
        Invoke-RemoveMode -CsvPath $InputCsvPath
    }
}

Write-Host ""
Write-StatusMessage "Script complete. Log: $log" -ForegroundColor Green
Write-LogEntry -LogName $log -LogEntryText "Script complete."

#endregion ────────────────────────────────────────────────────────────────────────────────────────
