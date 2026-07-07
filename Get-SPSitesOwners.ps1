<#
.SYNOPSIS
    Collects site owners for SharePoint Online sites in a tenant.

.DESCRIPTION
    This script connects to a SharePoint Online tenant and collects site owners.
    The information is exported to a CSV file for analysis.

.PARAMETER tenantname
    Your SharePoint Online tenant name (without .sharepoint.com)

.PARAMETER appID
    The Microsoft Entra (Azure AD) application ID for authentication

.PARAMETER thumbprint
    The certificate thumbprint for app-based authentication

.PARAMETER tenant
    Your tenant ID (GUID)

.PARAMETER inputfile
    Optional. Path to a CSV file containing a list of sites to process. If not provided, all sites will be processed.
    CSV should have a header of "URL" with site URLs in the first column.

.NOTES
    File Name      : Get-SPSitesOwners.ps1
    Author         : Mike Lee
    Prerequisite   : PnP.PowerShell module installed
    Date           : 11/13/25     
    Version        : 3.0

    Requirements:
        - PnP.PowerShell module installed
        - PowerShell 7.4 or higher
        - Appropriate permissions granted to the Azure AD application
            - SharePoint |Application | Sites.Read.All
            - Microsoft Graph| Application | Group.Read.All (for Entra group owners)
        - Certificate-based authentication configured
    
    The script collects owners from:
        - SharePoint Site Owners Group members
        - Site Collection Administrators
        - Entra (M365) Group Owners (for group-connected sites)

.EXAMPLE
    $tenantname = "contoso"
    $appID = "12345678-1234-1234-1234-1234567890ab"
    $thumbprint = "A1B2C3D4E5F6G7H8I9J0K1L2M3N4O5P6Q7R8S9T0"
    $tenant = "87654321-4321-4321-4321-ba0987654321"
    $inputfile = "C:\temp\sitelist-contoso.csv"
    .\Get-SPSitesOwners.ps1
#>

# Set Variables
$tenantname = "m365x61250205" #This is your tenant name
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"  #This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9" #This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51" #This is your Tenant ID

#Initialize Parameters - Do not change
$sites = @() # Array to hold site objects to be processed
$inputfile = $null # Path to the optional input CSV file for specific sites
$outputfile = $null # Path for the output CSV file
$log = $null # Path for the log file
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss" # Current date and time for unique file naming
$maxRetries = 5  # Maximum number of retry attempts for PnP cmdlets
$initialRetryDelay = 5  # Initial retry delay in seconds for PnP cmdlets

#Input / Output and Log Files
#$inputfile = "C:\temp\sitelist-m365x61250205.csv" # Example: This is the input file with list of sites to process. If not provided, all sites will be processed.
$outputfile = "$env:TEMP\" + 'SPSites_Owners_' + $date + '_' + "output.csv" # Define output CSV file path
$log = "$env:TEMP\" + 'SPSites_Owners_' + $date + '_' + "logfile.log" # Define log file path

#This is the logging function
Function Write-LogEntry {
    param(
        [string] $LogName, # Path to the log file
        [string] $LogEntryText, # Text to write to the log
        [string] $LogLevel = "INFO"  # Default log level is INFO (INFO, WARNING, ERROR)
    )
    if ($LogName -ne $null) {
        # log the date and time in the text file along with the data passed
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : [$LogLevel] $LogEntryText" | Out-File -FilePath $LogName -append;
    }
}


# Function to handle throttling with exponential backoff for PnP cmdlets
Function Invoke-PnPWithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock] $ScriptBlock, # The PnP command to execute
        
        [Parameter(Mandatory = $false)]
        [string] $Operation = "PnP Operation", # Description of the operation for logging
        
        [Parameter(Mandatory = $false)]
        [int] $MaxRetries = 5, # Maximum number of retries for this specific operation
        
        [Parameter(Mandatory = $false)]
        [int] $InitialRetryDelay = 5, # Initial delay in seconds before retrying
        
        [Parameter(Mandatory = $false)]
        [string] $LogName # Path to the log file
    )
    
    $retryCount = 0
    $success = $false
    $result = $null
    $retryDelay = $InitialRetryDelay
    
    do {
        try {
            # Execute the provided script block
            $result = & $ScriptBlock
            $success = $true
            return $result
        }
        catch {
            $exceptionDetails = $_.Exception.ToString()
            
            # Check for common throttling-related HTTP status codes or messages
            if (($exceptionDetails -like "*429*") -or 
                ($exceptionDetails -like "*throttl*") -or 
                ($exceptionDetails -like "*too many requests*") -or
                ($exceptionDetails -like "*request limit exceeded*")) {
                
                $retryCount++
                
                # Check if maximum retries have been reached
                if ($retryCount -ge $MaxRetries) {
                    Write-LogEntry -LogName $LogName -LogEntryText "Max retries ($MaxRetries) reached for $Operation. Giving up." -LogLevel "ERROR" 
                    throw $_ # Re-throw the original exception
                }
                
                # Parse Retry-After header from the exception response if available
                $retryAfterValue = $null
                if ($_.Exception.Response -and $_.Exception.Response.Headers -and $_.Exception.Response.Headers["Retry-After"]) {
                    $retryAfterValue = [int]$_.Exception.Response.Headers["Retry-After"]
                    $retryDelay = $retryAfterValue # Use server-suggested delay
                    Write-LogEntry -LogName $LogName -LogEntryText "Throttling detected for $Operation. Server requested retry after $retryAfterValue seconds." -LogLevel "WARNING"
                }
                else {
                    # Use exponential backoff if no Retry-After header is present
                    $retryDelay = [Math]::Min(60, $retryDelay * 2) # Double the delay, max 60 seconds
                    Write-LogEntry -LogName $LogName -LogEntryText "Throttling detected for $Operation. Using exponential backoff: waiting $retryDelay seconds before retry $retryCount of $MaxRetries." -LogLevel "WARNING"
                }
                
                Write-Host "Throttling detected for $Operation. Waiting $retryDelay seconds before retry $retryCount of $MaxRetries." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryDelay # Wait before retrying
            }
            else {
                # If not a throttling error, re-throw the original exception
                throw $_
            }
        }
    } while (-not $success -and $retryCount -lt $MaxRetries)
}

# Define the connection parameters for reuse across PnP cmdlets
$connectionParams = @{
    ClientId      = $appID         # Azure AD App ID for authentication
    Thumbprint    = $thumbprint    # Certificate thumbprint for app-based authentication
    Tenant        = $tenant         # Tenant ID (GUID)
    WarningAction = 'SilentlyContinue' # Suppress PnP warnings that are not errors
}

#Connect to SharePoint Admin Center initially
try {
    $adminUrl = 'https://' + $tenantname + '-admin.sharepoint.com' # Construct Admin Center URL
    
    # Connect using retry logic
    Invoke-PnPWithRetry -ScriptBlock { 
        Connect-PnPOnline -Url $adminUrl @connectionParams 
    } -Operation "Connect to SharePoint Admin Center" -LogName $Log
    
    Write-LogEntry -LogName $Log -LogEntryText "Successfully connected to SharePoint Admin Center: $adminUrl"
}
catch {
    # Handle connection failure
    Write-Host "Error connecting to SharePoint Admin Center ($adminUrl): $_" -ForegroundColor Red
    Write-LogEntry -LogName $Log -LogEntryText "Error connecting to SharePoint Admin Center ($adminUrl): $_" -LogLevel "ERROR"
    exit # Exit script if initial connection fails
}

# Get Site List: either from an input file or by querying all tenant sites
if ($inputfile -and (Test-Path -Path $inputfile)) {
    # Input file provided and exists
    try {
        $sites = Import-csv -path $inputfile -Header 'URL' # Import site URLs from CSV
        Write-LogEntry -LogName $Log -LogEntryText "Using sites from input file: $inputfile"
        Write-Host "Reading sites from input file: $inputfile" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Error reading input file '$inputfile': $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error reading input file '$inputfile': $_" -LogLevel "ERROR"
        exit # Exit if input file reading fails
    }
}
else {
    # No input file, or file not found; get all sites from the tenant
    Write-Host "Getting site list from tenant (this might take a while)..." -ForegroundColor Yellow
    Write-LogEntry -LogName $Log -LogEntryText "Getting sites using Get-PnPTenantSite (no input file specified or found)"
    try {
        # Ensure connection to Admin Center before getting tenant sites
        Invoke-PnPWithRetry -ScriptBlock { 
            Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop 
        } -Operation "Connect to SharePoint Admin Center (before Get-PnPTenantSite)" -LogName $Log
        
        # Retrieve all tenant sites, excluding MySites and RedirectSites
        $sites = Invoke-PnPWithRetry -ScriptBlock { 
            Get-PnPTenantSite | Where-Object { $_.Url -notlike "*-my.sharepoint.com*" -and $_.Template -ne "RedirectSite#0" }
        } -Operation "Get-PnPTenantSite" -LogName $Log
        
        Write-Host "Found $($sites.Count) sites." -ForegroundColor Green
        Write-LogEntry -LogName $Log -LogEntryText "Retrieved $($sites.Count) sites using Get-PnPTenantSite."
    }
    catch {
        Write-Host "Error getting site list from tenant: $_" -ForegroundColor Red
        Write-LogEntry -LogName $Log -LogEntryText "Error getting site list from tenant: $_" -LogLevel "ERROR"
        exit # Exit if fetching all sites fails
    }
}


$totalSites = $sites.Count # Total number of sites to process
$processedCount = 0 # Counter for processed sites

# Define CSV Headers for the output file
$csvHeaders = "URL,SharePoint Site Owners,Site Collection Admins,Entra Group Owners"

# Create the output CSV file and write the headers
Set-Content -Path $outputfile -Value $csvHeaders -Encoding UTF8
Write-Host "Created output file with headers: $outputfile" -ForegroundColor Green
Write-LogEntry -LogName $Log -LogEntryText "Created output file with headers: $outputfile"

# Main processing loop: Iterate through each site
foreach ($site in $sites) {
    $processedCount++
    $siteUrl = $site.Url 
    Write-Host "Processing site $processedCount/$totalSites : $siteUrl" -ForegroundColor Cyan
    Write-LogEntry -LogName $Log -LogEntryText "Processing site $processedCount/$totalSites : $siteUrl"

    # Initialize owner collections
    $spSiteOwners = @()
    $siteAdmins = @()
    $entraOwners = @()

    try {
        # Connect to Admin URL to get tenant-level properties for the site
        Invoke-PnPWithRetry -ScriptBlock { 
            Connect-PnPOnline -Url $adminUrl @connectionParams -ErrorAction Stop 
        } -Operation "Connect to Admin URL for site props $siteUrl" -LogName $Log
        
        # Get tenant-level site properties (URL and GroupId)
        $siteprops = Invoke-PnPWithRetry -ScriptBlock { 
            Get-PnPTenantSite -Identity $siteUrl | Select-Object Url, GroupId
        } -Operation "Get-PnPTenantSite for $siteUrl" -LogName $Log

        # If site properties couldn't be retrieved, log error and skip to the next site
        if ($null -eq $siteprops) { 
            Write-LogEntry -LogName $Log -LogEntryText "Failed to retrieve properties for site $siteUrl. Skipping." -LogLevel "ERROR"
            continue 
        }

        # Connect to the specific site to get owner information
        try {
            Invoke-PnPWithRetry -ScriptBlock { 
                Connect-PnPOnline -Url $siteUrl @connectionParams -ErrorAction Stop 
            } -Operation "Connect to site $siteUrl" -LogName $Log

            # Get SharePoint Site Owners Group members
            try {
                # Try to get the Associated Owner Group first (most reliable method)
                $ownersGroup = $null
                
                try {
                    $web = Invoke-PnPWithRetry -ScriptBlock { 
                        Get-PnPWeb -Includes AssociatedOwnerGroup
                    } -Operation "Get Web with Associated Owner Group for $siteUrl" -LogName $Log
                    
                    if ($web.AssociatedOwnerGroup) {
                        $ownersGroup = $web.AssociatedOwnerGroup
                        Write-LogEntry -LogName $Log -LogEntryText "Found Associated Owner Group: $($ownersGroup.Title) for $siteUrl"
                    }
                }
                catch {
                    Write-LogEntry -LogName $Log -LogEntryText "Could not retrieve AssociatedOwnerGroup for $siteUrl, will try alternate methods" -LogLevel "WARNING"
                }
                
                # Fallback: If AssociatedOwnerGroup didn't work, look for groups with "Owners" in the name
                if (-not $ownersGroup) {
                    $ownersGroup = Invoke-PnPWithRetry -ScriptBlock { 
                        Get-PnPGroup | Where-Object { $_.Title -like "*Owners" -or $_.Title -like "*owners" }
                    } -Operation "Get Owners Group by name for $siteUrl" -LogName $Log
                    
                    if ($ownersGroup -and $ownersGroup.Count -gt 1) {
                        # If multiple groups found, take the first one
                        $ownersGroup = $ownersGroup[0]
                    }
                }
                
                if ($ownersGroup) {
                    Write-LogEntry -LogName $Log -LogEntryText "Using Owners Group: '$($ownersGroup.Title)' (ID: $($ownersGroup.Id)) for $siteUrl"
                    
                    # Get members of the Owners group
                    $ownersGroupMembers = Invoke-PnPWithRetry -ScriptBlock { 
                        Get-PnPGroupMember -Identity $ownersGroup.Id
                    } -Operation "Get Owners Group Members for $siteUrl" -LogName $Log
                    
                    if ($ownersGroupMembers) {
                        foreach ($member in $ownersGroupMembers) {
                            # Skip groups nested within the owners group, only get users
                            if ($member.PrincipalType -eq "User" -or $member.PrincipalType -eq 1) {
                                if ($member.Email) {
                                    $spSiteOwners += "$($member.Title) <$($member.Email)>"
                                }
                                elseif ($member.LoginName) {
                                    $spSiteOwners += "$($member.Title) ($($member.LoginName))"
                                }
                                else {
                                    $spSiteOwners += $member.Title
                                }
                            }
                        }
                        Write-LogEntry -LogName $Log -LogEntryText "Found $($spSiteOwners.Count) user members in SharePoint Owners group '$($ownersGroup.Title)' for $siteUrl"
                    }
                    else {
                        Write-LogEntry -LogName $Log -LogEntryText "Owners group '$($ownersGroup.Title)' exists but has no members for $siteUrl" -LogLevel "WARNING"
                    }
                }
                else {
                    Write-LogEntry -LogName $Log -LogEntryText "No Owners group found for $siteUrl using any method" -LogLevel "WARNING"
                }
            }
            catch {
                Write-LogEntry -LogName $Log -LogEntryText "Error retrieving SharePoint Owners group members for $siteUrl : $_" -LogLevel "WARNING"
            }

            # Get Site Collection Administrators
            try {
                $siteCollAdmins = Invoke-PnPWithRetry -ScriptBlock { 
                    Get-PnPSiteCollectionAdmin 
                } -Operation "Get-PnPSiteCollectionAdmin for $siteUrl" -LogName $Log
                
                foreach ($admin in $siteCollAdmins) {
                    if ($admin -and $admin.Email) {
                        $siteAdmins += "$($admin.Title) <$($admin.Email)>"
                    }
                    elseif ($admin -and $admin.Title) {
                        $siteAdmins += $admin.Title
                    }
                }
                Write-LogEntry -LogName $Log -LogEntryText "Found $($siteAdmins.Count) site collection admins for $siteUrl"
            }
            catch {
                Write-LogEntry -LogName $Log -LogEntryText "Error retrieving site collection admins for $siteUrl : $_" -LogLevel "WARNING"
            }

            # If the site is Microsoft 365 Group-connected, get Entra Group Owners
            if ($null -ne $siteprops.GroupId -and $siteprops.GroupId -ne [System.Guid]::Empty) {
                try {
                    Write-LogEntry -LogName $Log -LogEntryText "Site $siteUrl is group-connected. GroupId: $($siteprops.GroupId)"
                    
                    # Get M365 Group Owners
                    $groupOwners = Invoke-PnPWithRetry -ScriptBlock { 
                        Get-PnPMicrosoft365GroupOwners -Identity $siteprops.GroupId 
                    } -Operation "Get M365 Group Owners for $($siteprops.GroupId)" -LogName $Log
                    
                    foreach ($owner in $groupOwners) {
                        if ($owner.Mail) {
                            $entraOwners += "$($owner.DisplayName) <$($owner.Mail)>"
                        }
                        elseif ($owner.UserPrincipalName) {
                            $entraOwners += "$($owner.DisplayName) <$($owner.UserPrincipalName)>"
                        }
                        else {
                            $entraOwners += $owner.DisplayName
                        }
                    }
                    Write-LogEntry -LogName $Log -LogEntryText "Found $($entraOwners.Count) Entra group owners for $siteUrl"
                }
                catch {
                    Write-LogEntry -LogName $Log -LogEntryText "Warning: Could not retrieve M365 group owners for $($siteprops.GroupId) on $siteUrl : $_" -LogLevel "WARNING"
                }
            }
        }
        catch {
            Write-LogEntry -LogName $Log -LogEntryText "Could not connect to site $siteUrl to get additional owner info. $_" -LogLevel "WARNING"
        }

        # Create output object with all owner information
        $exportItem = [PSCustomObject]@{
            URL                      = $siteprops.Url
            "SharePoint Site Owners" = ($spSiteOwners -join '; ')
            "Site Collection Admins" = ($siteAdmins -join '; ')
            "Entra Group Owners"     = ($entraOwners -join '; ')
        }

        # Export to CSV
        $exportItem | Export-Csv -Path $outputfile -NoTypeInformation -Append -Encoding UTF8
        Write-Host "Exported data for site $processedCount/$totalSites to CSV" -ForegroundColor Green
        Write-LogEntry -LogName $Log -LogEntryText "Successfully wrote data for site $siteUrl to CSV"
    }
    catch {
        Write-LogEntry -LogName $Log -LogEntryText "ERROR: Could not process site $siteUrl. $_" -LogLevel "ERROR"
        continue
    }
} # End foreach Site

# Disconnect PnP Online session if one exists
if (Get-PnPConnection) {
    Disconnect-PnPOnline
}
Write-LogEntry -LogName $Log -LogEntryText "Disconnected from PnP Online. Script finished."
Write-Host "Script finished. Log file located at: $log" -ForegroundColor Green
Write-Host "Output CSV located at: $outputfile" -ForegroundColor Green
