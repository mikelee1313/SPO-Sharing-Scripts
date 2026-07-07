<#
.SYNOPSIS
Locates SharePoint Remote Event Receivers across the tenant in preparation for MC1411726.

.DESCRIPTION
This script connects to SharePoint Online by using app-only certificate authentication,
enumerates tenant sites, inspects site collection, web, and non-hidden list event
receivers across each root web and subweb, and records Remote Event Receivers that
should be reviewed for the MC1411726 retirement effort.

The script excludes personal OneDrive sites and excludes
Microsoft.SharePoint.Webhooks.SPWebhookItemEventReceiver because webhook receivers are
not part of this retirement scope.

.OUTPUTS
Creates two files in the current user's temp folder:
- A log file that records scan progress, discovered receivers, and errors.
- A CSV file that contains the receiver inventory.

The CSV includes these columns:
SiteUrl, WebUrl, Scope, ListName, ReceiverName, ReceiverClass, ReceiverAssembly, ReceiverUrl, EventType, Type.

.NOTES
The app registration used by this script must be able to enumerate tenant sites and
read SharePoint list event receiver information.

Required API permission:
- SharePoint application permission `Sites.FullControl.All` with admin consent.

This script uses PnP PowerShell against SharePoint Online, so Microsoft Graph
permissions are not required for this specific inventory.

.AUTHOR
Mike Lee
Date: 7/7/2026
#>

$appID = "abc64618-283f-47ba-a185-50d935d51d57"  #This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9" #This is certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3" #This is your Tenant ID
$adminUrl = "https://m365cpi13246019-admin.sharepoint.com"
$logPath = "$env:TEMP\MC1411726-RemoteEventReceivers-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss")
$csvPath = "$env:TEMP\MC1411726-RemoteEventReceivers-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss")

# Define the connection parameters for reuse across PnP cmdlets
$connectionParams = @{
    ClientId      = $appID         # Azure AD App ID for authentication
    Thumbprint    = $thumbprint    # Certificate thumbprint for app-based authentication
    Tenant        = $tenant         # Tenant ID (GUID)
    WarningAction = 'SilentlyContinue' # Suppress PnP warnings that are not errors
}

function Get-AffectedRemoteEventReceivers
{
    param(
        [AllowNull()]
        [object[]]$EventReceivers
    )

    if ($null -eq $EventReceivers)
    {
        return @()
    }

    @(
        $EventReceivers |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.ReceiverUrl) -and
            $_.ReceiverClass -notlike '*SPWebhook*'
        }
    )
}

function Add-ReceiverInventoryRows
{
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Results,

        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$WebUrl,

        [Parameter(Mandatory = $true)]
        [string]$Scope,

        [string]$ListName,

        [Parameter(Mandatory = $true)]
        [object[]]$Receivers,

        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    foreach ($receiver in $Receivers)
    {
        $Results.Add([pscustomobject]@{
            SiteUrl          = $SiteUrl
            WebUrl           = $WebUrl
            Scope            = $Scope
            ListName         = $ListName
            ReceiverName     = $receiver.ReceiverName
            ReceiverClass    = $receiver.ReceiverClass
            ReceiverAssembly = $receiver.ReceiverAssembly
            ReceiverUrl      = $receiver.ReceiverUrl
            EventType        = $receiver.EventType
            Type             = $receiver.Type
        })

        if ($Scope -eq 'List')
        {
            Add-Content -Path $LogPath -Value ("[{0}] Found {1} scope receiver {2} on list {3} in web {4}" -f (Get-Date -Format s), $Scope, $receiver.ReceiverName, $ListName, $WebUrl)
        }
        else
        {
            Add-Content -Path $LogPath -Value ("[{0}] Found {1} scope receiver {2} in web {3}" -f (Get-Date -Format s), $Scope, $receiver.ReceiverName, $WebUrl)
        }
    }
}

function Get-VisibleListsForCurrentWeb
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$WebUrl,

        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    try
    {
        return @(
            Get-PnPList |
            Where-Object { -not $_.Hidden }
        )
    }
    catch
    {
        Add-Content -Path $LogPath -Value ("[{0}] Get-PnPList failed for web {1}. Falling back to Get-PnPWeb/Get-PnPProperty. Error: {2}" -f (Get-Date -Format s), $WebUrl, $_.Exception.Message)

        $web = Get-PnPWeb
        Get-PnPProperty -ClientObject $web -Property Lists | Out-Null

        return @(
            $web.Lists |
            Where-Object { -not $_.Hidden }
        )
    }
}

Connect-PnPOnline -Url $adminUrl @connectionParams

$sites = Get-PnPTenantSite | Where-Object { $_.Url -notlike "*-my.sharepoint.com*" -and $_.Template -ne "RedirectSite#0" -and $_.ArchiveStatus -eq "NotArchived" }
$siteCount = $sites.Count
$siteIndex = 0
$totalReceiverCount = 0

Add-Content -Path $logPath -Value ("[{0}] Starting scan across {1} sites" -f (Get-Date -Format s), $siteCount)
Write-Host ("Starting scan across {0} SharePoint site(s)" -f $siteCount) -ForegroundColor Cyan

foreach ($site in $sites)
{
    $siteIndex++
    $sitesRemaining = $siteCount - $siteIndex
    $progressStatus = "{0} ({1}/{2}) - {3} remaining" -f $site.Url, $siteIndex, $siteCount, $sitesRemaining
    Write-Progress -Activity "Scanning SharePoint sites for remote event receivers" -Status $progressStatus -PercentComplete (($siteIndex / $siteCount) * 100)
    Add-Content -Path $logPath -Value ("[{0}] Processing site {1} ({2}/{3})" -f (Get-Date -Format s), $site.Url, $siteIndex, $siteCount)

    try
    {
        $currentOperation = 'Connect-PnPOnline to site'
        Connect-PnPOnline -Url $site.Url @connectionParams
        $siteResults = New-Object System.Collections.Generic.List[object]
        $currentOperation = 'Get-PnPEventReceiver -Scope Site'
        $siteLevelReceivers = Get-AffectedRemoteEventReceivers -EventReceivers (Get-PnPEventReceiver -Scope Site)

        if ($siteLevelReceivers.Count -gt 0)
        {
            Add-ReceiverInventoryRows -Results $siteResults -SiteUrl $site.Url -WebUrl $site.Url -Scope 'Site' -ListName '' -Receivers $siteLevelReceivers -LogPath $logPath
        }

        $webUrls = @($site.Url)
    $currentOperation = 'Get-PnPSubWeb -Recurse'
        $subWebUrls = @(Get-PnPSubWeb -Recurse | Select-Object -ExpandProperty Url)

        if ($subWebUrls.Count -gt 0)
        {
            $webUrls += $subWebUrls
        }

        foreach ($webUrl in $webUrls)
        {
            try
            {
                $currentOperation = 'Connect-PnPOnline to web'
                Connect-PnPOnline -Url $webUrl @connectionParams
                $currentOperation = 'Get-PnPEventReceiver -Scope Web'
                $webLevelReceivers = Get-AffectedRemoteEventReceivers -EventReceivers (Get-PnPEventReceiver -Scope Web)

                if ($webLevelReceivers.Count -gt 0)
                {
                    Add-ReceiverInventoryRows -Results $siteResults -SiteUrl $site.Url -WebUrl $webUrl -Scope 'Web' -ListName '' -Receivers $webLevelReceivers -LogPath $logPath
                }

                $currentOperation = 'Get list collection'
                $lists = Get-VisibleListsForCurrentWeb -WebUrl $webUrl -LogPath $logPath

                foreach ($list in $lists)
                {
                    $currentOperation = "Get-PnPEventReceiver -List $($list.Title)"
                    $listLevelReceivers = Get-AffectedRemoteEventReceivers -EventReceivers (Get-PnPEventReceiver -List $list)

                    if ($listLevelReceivers.Count -gt 0)
                    {
                        Add-ReceiverInventoryRows -Results $siteResults -SiteUrl $site.Url -WebUrl $webUrl -Scope 'List' -ListName $list.Title -Receivers $listLevelReceivers -LogPath $logPath
                    }
                }
            }
            catch
            {
                Add-Content -Path $logPath -Value ("[{0}] Failed to process web {1} in site {2} during {3}: {4}" -f (Get-Date -Format s), $webUrl, $site.Url, $currentOperation, $_.Exception.Message)
            }

        }

        if ($siteResults.Count -gt 0)
        {
            $siteResults | Export-Csv -Path $csvPath -NoTypeInformation -Append
            $totalReceiverCount += $siteResults.Count
            Add-Content -Path $logPath -Value ("[{0}] Wrote {1} receiver records for site {2} to CSV" -f (Get-Date -Format s), $siteResults.Count, $site.Url)
            Write-Host ("Completed site {0}/{1} ({2} remaining): {3} found {4} Remote Event Receiver(s) affected by MC1411726" -f $siteIndex, $siteCount, $sitesRemaining, $site.Url, $siteResults.Count) -ForegroundColor Red
        }
        else
        {
            Add-Content -Path $logPath -Value ("[{0}] No Remote Event Receivers affected by MC1411726 found in site {1}" -f (Get-Date -Format s), $site.Url)
            Write-Host ("Completed site {0}/{1} ({2} remaining): {3} found 0 Remote Event Receivers affected by MC1411726" -f $siteIndex, $siteCount, $sitesRemaining, $site.Url) -ForegroundColor Green
        }
    }
    catch
    {
        Add-Content -Path $logPath -Value ("[{0}] Failed to process site {1} during {2}: {3}" -f (Get-Date -Format s), $site.Url, $currentOperation, $_.Exception.Message)
    }
}

Write-Progress -Activity "Scanning SharePoint sites for remote event receivers" -Completed

if ($totalReceiverCount -eq 0)
{
    Add-Content -Path $logPath -Value ("[{0}] Scan complete. No Remote Event Receivers affected by MC1411726 were found." -f (Get-Date -Format s))
    Write-Host "No Remote Event Receivers affected by MC1411726 were found." -ForegroundColor Green
    Write-Host ("Remote event receiver log written to {0}" -f $logPath) -ForegroundColor Green
}
else
{
    Add-Content -Path $logPath -Value ("[{0}] Scan complete. Found {1} Remote Event Receivers affected by MC1411726. CSV output written to {2}" -f (Get-Date -Format s), $totalReceiverCount, $csvPath)
    Write-Host ("Remote event receiver log written to {0}" -f $logPath) -ForegroundColor Yellow
    Write-Host ("Remote event receiver CSV written to {0}" -f $csvPath) -ForegroundColor Yellow
}
