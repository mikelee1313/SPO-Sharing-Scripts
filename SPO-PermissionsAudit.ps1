<#
.SYNOPSIS
    Audits SharePoint Online permission changes across a list of site collections
    using the Unified Audit Log.

.DESCRIPTION
    Queries Search-UnifiedAuditLog for SPO permission events:
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

.PARAMETER SiteListPath
    Path to a text file containing one SPO site URL per line.

.PARAMETER StartDate
    Start of the audit window (UTC). Defaults to 90 days ago.

.PARAMETER EndDate
    End of the audit window (UTC). Defaults to now.

.PARAMETER OutputPath
    Path for the exported CSV. Defaults to .\SPO-PermissionsAudit_<timestamp>.csv

.PARAMETER ResultSize
    Max records returned per Search-UnifiedAuditLog call. Max allowed is 5000.

.EXAMPLE
    .\SPO-PermissionsAudit.ps1 -SiteListPath .\sites.txt -StartDate (Get-Date).AddDays(-30)

.NOTES
    Author  : Mike Lee / Mariel Williams
    Created : 3/27/2026
    Updated : 3/31/2026 - added retry logic for EXO throttling, improved output formatting, and added checkpointing for long runs
    Version : 1.1
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$SiteListPath = "C:\temp\SPOSiteList.txt",

    [Parameter()]
    [datetime]$StartDate = (Get-Date).ToUniversalTime().AddDays(-3),

    [Parameter()]
    [datetime]$EndDate = (Get-Date).ToUniversalTime(),

    [Parameter()]
    [string]$OutputPath = ".\SPO-PermissionsAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter()]
    [ValidateRange(1, 5000)]
    [int]$ResultSize = 5000,

    # Suppress internal SPO system-group rows (auto-generated Limited Access side-effects)
    [Parameter()]
    [switch]$IncludeSystemEvents
)

#region ── Prerequisites ────────────────────────────────────────────────────────

# Verify ExchangeOnlineManagement is available (provides Search-UnifiedAuditLog)
if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

#endregion

#region ── Configuration ────────────────────────────────────────────────────────



# Unified Audit Log operations that represent SPO permission changes
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

function Get-AuditEventsForSite {
    <#
    .SYNOPSIS  Pages through Search-UnifiedAuditLog for a single site URL.
    #>
    param (
        [string]   $SiteUrl,
        [datetime] $Start,
        [datetime] $End,
        [string[]] $Operations,
        [int]      $PageSize
    )

    $allRecords = [System.Collections.Generic.List[PSObject]]::new()
    $sessionId = "SPOPermAudit-$(New-Guid)"
    $page = 1

    Write-Verbose "  Querying UAL for: $SiteUrl"

    do {
        # Retry with exponential back-off to handle EXO throttling
        $attempt = 0
        $results = $null
        do {
            try {
                $results = Search-UnifiedAuditLog `
                    -StartDate       $Start `
                    -EndDate         $End `
                    -Operations      $Operations `
                    -ObjectIds       "$SiteUrl*" `
                    -SessionId       $sessionId `
                    -SessionCommand  ReturnLargeSet `
                    -ResultSize      $PageSize `
                    -ErrorAction     Stop
                break  # success - exit retry loop
            }
            catch {
                $attempt++
                if ($attempt -ge 3) { throw }
                $delay = [math]::Pow(2, $attempt) * 5  # 10s, 20s
                Write-Warning "    UAL query failed (attempt $attempt/3), retrying in ${delay}s: $_"
                Start-Sleep -Seconds $delay
            }
        } while ($true)

        if ($results) {
            foreach ($r in $results) { $allRecords.Add($r) }
            Write-Verbose "    Page $page - retrieved $($results.Count) records (running total: $($allRecords.Count))"
            $page++
        }
    } while ($results -and $results.Count -eq $PageSize)

    return $allRecords
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
    .SYNOPSIS  Flattens a UAL record into a clean, admin-readable object.
    #>
    param ([PSObject]$Record)

    try {
        $audit = $Record.AuditData | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $audit = $null
    }

    $op = $Record.Operations
    $action = if ($script:ActionLabels[$op]) { $script:ActionLabels[$op] } else { $op }

    # Parse PermissionsGranted and GroupAffected out of the EventData XML blob
    $permGranted = ''
    $groupName = ''
    if ($audit.EventData) {
        if ($audit.EventData -match '<PermissionsGranted>([^<]+)<') { $permGranted = $Matches[1] }
        if ($audit.EventData -match '<Group>([^<]+)<') { $groupName = $Matches[1] }
    }

    # Relative path: SourceRelativeUrl is cleanest; fall back to stripping the site URL from ObjectId
    $relPath = ''
    if ($audit.SourceRelativeUrl) {
        $relPath = $audit.SourceRelativeUrl
    }
    elseif ($audit.ObjectId -and $audit.SiteUrl) {
        $stripped = $audit.ObjectId -replace [regex]::Escape($audit.SiteUrl.TrimEnd('/')), ''
        $relPath = if ($stripped -match '^[/\\]?$') { '(site root)' } else { $stripped.TrimStart('/') }
    }

    # Clean up target name — internal SharingLinks group names are not meaningful to admins
    $target = if ($audit.TargetUserOrGroupName -match '^SharingLinks\.') {
        '(sharing link group)'
    }
    else {
        $audit.TargetUserOrGroupName
    }

    # Flag system-generated side-effect rows so they can be filtered
    $isSystem = ($op -eq 'AddedToGroup') -and (
        $groupName -match '^Limited Access System Group' -or
        $groupName -match '^SharingLinks\.'
    )

    # Friendly link scope (blank when not applicable)
    $linkScope = if ($audit.SharingLinkScope -and $audit.SharingLinkScope -notin 'Uninitialized', 'None') {
        $audit.SharingLinkScope
    }
    else { '' }

    [PSCustomObject]@{
        DateTime           = $Record.CreationDate
        PerformedBy        = $Record.UserIds
        Action             = $action
        ItemType           = $audit.ItemType
        RelativePath       = $relPath
        SiteUrl            = $audit.SiteUrl
        AffectedUser       = $target
        PermissionsGranted = $permGranted
        GroupAffected      = $groupName
        LinkScope          = $linkScope
        ClientIP           = $audit.ClientIP
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

Write-Host "SPO Permissions Audit" -ForegroundColor Cyan
Write-Host "  Sites    : $($sites.Count)" -ForegroundColor Cyan
Write-Host "  Window   : $($StartDate.ToString('u'))  - >  $($EndDate.ToString('u'))" -ForegroundColor Cyan
Write-Host "  Operations: $($PermissionOperations.Count) event types" -ForegroundColor Cyan
Write-Host ""

# Checkpoint file tracks completed sites so the run can resume after a failure
$checkpointFile = "$OutputPath.checkpoint"
$completedSites = [System.Collections.Generic.HashSet[string]]::new()
if (Test-Path $checkpointFile) {
    Write-Host "  Resuming from checkpoint: $checkpointFile" -ForegroundColor Yellow
    Get-Content $checkpointFile | Where-Object { $_ -ne '' } | ForEach-Object { $completedSites.Add($_) | Out-Null }
}

$pendingSites = $sites | Where-Object { -not $completedSites.Contains($_) }

if ($completedSites.Count -gt 0) {
    Write-Host "  Skipping  : $($completedSites.Count) already-processed site(s)" -ForegroundColor DarkGray
}

$totalWritten = 0
$totalFiltered = 0
$siteIndex = $completedSites.Count  # start progress counter from where we left off

foreach ($site in $pendingSites) {
    $siteIndex++

    # Refresh EXO session every 50 sites to prevent token expiry on long runs
    if (($siteIndex % 50) -eq 0) {
        Write-Verbose "  Refreshing Exchange Online connection..."
        Connect-ExchangeOnline -ShowBanner:$false
    }

    Write-Progress -Activity "Querying Unified Audit Log" `
        -Status       "[$siteIndex/$($sites.Count)] $site" `
        -PercentComplete (($siteIndex / $sites.Count) * 100)

    try {
        $records = Get-AuditEventsForSite `
            -SiteUrl    $site `
            -Start      $StartDate `
            -End        $EndDate `
            -Operations $PermissionOperations `
            -PageSize   $ResultSize

        if ($records -and $records.Count -gt 0) {
            foreach ($rec in $records) {
                $flat = ConvertTo-FlatRecord -Record $rec
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
            Write-Host "  [+] $site  -  $($records.Count) event(s)" -ForegroundColor Green
        }
        else {
            Write-Host "  [ ] $site  -  no events found" -ForegroundColor Gray
        }

        # Mark site as done in checkpoint file
        Add-Content -Path $checkpointFile -Value $site
        [void]$completedSites.Add($site)
    }
    catch {
        Write-Warning "  [!] Error querying '$site': $_"
    }
}

Write-Progress -Activity "Querying Unified Audit Log" -Completed

#endregion

#region ── Output ───────────────────────────────────────────────────────────────

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

    # Remove checkpoint on clean completion
    if (Test-Path $checkpointFile) {
        Remove-Item $checkpointFile -Force
        Write-Host "`nCheckpoint cleared." -ForegroundColor DarkGray
    }
}

#endregion
