<#   
.SYNOPSIS
    Get-SPOSiteOwnersandSegments.ps1 - Loops through all sites and exports all groups and users for each site.

.DESCRIPTION
    This script connects to SharePoint Online and Exchange Online services to retrieve information about sites and their associated groups and users. It exports the collected data to a CSV file.

.PARAMETER Tenant
    Specifies the name of the tenant to connect to.

.PARAMETER Admin
    Specifies the admin account to use for site collection administration.

.INPUTS
    The script requires a CSV file containing a list of site URLs.

.OUTPUTS
    The script exports the collected site information, including groups and users, to a CSV file.

.NOTES
    Authors: Mike Lee, Kiran Bellala, Brian Mokaya
    Date: 6/11/2024
    Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
#>


#Configurable Settings
$t = 'M365x13453069' # < - Your Tenant Name Here
$admin = 'admin@M365x13453069.onmicrosoft.com'  # <- Your Admin Account Here

#Initialize Parameters - Do not change
$sites = @()
$output = @()
$inputfile = @()
$outputfile = @()
$date = @()
$log = @()
$date = Get-Date -Format yyyy-MM-dd_HH-mm-ss

#Input / Output and Log Files
$inputfile = 'C:\temp\sitelist.csv'
$outputfile = "$env:TEMP\" + 'Site_Entra_Group_Owners_' + $date + '_' + "output.csv"
$log = "$env:TEMP\" + 'Site_Entra_Group_Owners_' + $date + '_' + "logfile.log"

#Connect to Services
Connect-SPOService -Url ('https://' + $t + '-admin.sharepoint.com')
Connect-ExchangeOnline


#This is the logging function
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


#Get All Sites that are not Group Connected or Personal
#$sites = get-sposite -Limit All -IncludePersonalSite:$false -GroupIdDefined:$false  -Filter { 'Url' -notlike '-my.sharepoint.com' } | where { $_.Template -ne 'RedirectSite#0' | where $_.url -notcontains 'sharepoint.com/portals/' }

#All ShaerPoint Sites:
#$sites = get-sposite -Limit All | where { $_.Template -ne 'RedirectSite#0'}

#OneDrive Sites:
#$sites = get-sposite -Limit All -IncludePersonalSite:$true | where { $_.Template -like '*SPSPERS#*'}

#Get All Sites from a list

#Use Export from SP Admin
#$sites = Import-csv -path $inputfile -Header ('"Site name"','URL','Teams','Channel sites','Storage used (GB)','Primary admin','Hub', 'Template','Last activity (UTC)','Date created','Created by','Storage limit (GB)','Storage used (%)','Microsoft 365 group','Files viewed or edited','Page views','Page visits','Files','Sensitivity','External sharing', 'Segments') | Select-Object -Skip 1

#use simple imput file
$sites = Import-csv -path $inputfile -Header 'URL'

#Add account as admin and export groups
foreach ($site in $sites) {

    $AADGroups = @()
    $groupowners = @()
    $groupmembers = @()

    Write-Host "Starting Site enumeration..."
    Write-LogEntry -LogName:$Log -LogEntryText "Starting Site enumeration..."

    Write-Host "Site URL: $($site.Url)" -ForegroundColor Magenta
    Write-LogEntry -LogName:$Log -LogEntryText "Site URL: $($site.Url)"

    #Setting Admin Account as a Site Collection Admin
    try { 
        Write-Host "Attempting to SET $admin to  '$($site.url)' as Site Admin" -ForegroundColor Yellow
        Write-LogEntry -LogName:$Log -LogEntryText "Attempting to set '$admin' to  '$($site.url)' as Site Admin"
        $Addadmin = Set-SPOUser -Site $site.Url -LoginName $admin -IsSiteCollectionAdmin $true
        #sleep 1
    }        

    catch {
        Write-Host "Unable to Add '$admin' to  '$($site.url)' as Site Admin" -ForegroundColor Red
        Write-LogEntry -LogName:$Log -LogEntryText "Unable to Add '$admin' to '$($site.url)' as Site Admin"
    }
 
    #Get SPO Site Information Barrier Modes

    $siteprops = get-sposite -Identity $site.url | Select-Object URL, Owner, InformationBarriersMode, InformationSegment, GroupId, RelatedGroupId, IsHubSite, Template, SiteDefinedSharingCapability, SharingCapability, DisableCompanyWideSharingLinks, IsTeamsConnected, IsTeamsChannelConnected, TeamsChannelType

    if ($siteprops.GroupId.Guid -ne '00000000-0000-0000-0000-000000000000') {
        $groupowners = Get-UnifiedGroupLinks -Identity  $siteprops.GroupId -LinkType Owners
        $groupmembers = Get-UnifiedGroupLinks -Identity  $siteprops.GroupId -LinkType Members
      
        $gowner = @()
        $gmember = @()
        # This script block iterates over each owner in a group.
        # For each owner, it attempts to retrieve the owner's display name, primary SMTP address, and information barrier segments using the Get-Recipient cmdlet.
        # The retrieved information is then logged using a custom Write-LogEntry function.
        # If the script is unable to retrieve the information for an owner, it prints an error message to the console and logs the error.

        foreach ($groupowner in $groupowners) {
            try {
                # Attempt to retrieve owner information
                $gowner = Get-Recipient -Identity $groupowner.PrimarySmtpAddress | Select-Object DisplayName, PrimarySmtpAddress, InformationBarrierSegments
                # Log owner information  
                Write-LogEntry -LogName:$Log "Entra Group Owner is" $($gowner.DisplayName)
                Write-LogEntry -LogName:$Log "Entra Group Owner E-Mail Address is" $($gowner.PrimarySmtpAddress)
                Write-LogEntry -LogName:$Log "Entra Groups Onwer InfoSegment is" $($gowner.InformationBarrierSegments)

                       
                $ExportItem = New-Object PSObject
                $ExportItem  | Add-Member -MemberType NoteProperty -name "URL" -value ($($siteprops.url))
                $ExportItem  | Add-Member -MemberType NoteProperty -name "Owner" -value ($($siteprops.Owner))  
                $ExportItem  | Add-Member -MemberType NoteProperty -name "IB Mode" -value ($($siteprops.InformationBarriersMode) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "IB Segment" -value ($($siteprops.InformationSegment) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "Group ID" -value ($($siteprops.GroupId) -join ',')    
                $ExportItem  | Add-Member -MemberType NoteProperty -name "RelatedGroupId" -value ($($siteprops.RelatedGroupId) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "IsHubSite" -value ($($siteprops.IsHubSite) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "Template" -value ($($siteprops.Template) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "SiteDefinedSharingCapability" -value ($($siteprops.SiteDefinedSharingCapability) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "SharingCapability" -value ($($siteprops.SharingCapability) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "DisableCompanyWideSharingLinks" -value ($($siteprops.DisableCompanyWideSharingLinks) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "IsTeamsConnected" -value ($($siteprops.IsTeamsConnected) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "IsTeamsChannelConnected" -value ($($siteprops.IsTeamsChannelConnected) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "TeamsChannelType" -value ($($siteprops.TeamsChannelType) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "Entra Group Owners" -value ($($gowner.DisplayName) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "Entra Group Owners Email" -value ($($gowner.PrimarySmtpAddress) -join ',')
                $ExportItem  | Add-Member -MemberType NoteProperty -name "Entra Groups Owners InfoSegment" -value ($($gowner.InformationBarrierSegments) -join ',')
                $output += $ExportItem

            }        
            catch {
                # Print and log error message if unable to retrieve owner information
                Write-Host "Unable to retrieve information for group owner: $groupowner" -ForegroundColor Red
                Write-LogEntry -LogName:$Log -LogEntryText "Unable to retrieve information for group owner: $groupowner"
            }

        }   
  
    }

    #Removing Admin Account as a Site Collection Admin
    Try {
       
        Write-Host "Attempting to Remove $admin to  '$($site.url)' as Site Admin" -ForegroundColor Yellow
        Write-LogEntry -LogName:$Log -LogEntryText "Attempting to Remove '$admin' to  '$($site.url)' as Site Admin"

        $removeadmin = Set-SPOUser -Site $site.Url -LoginName $admin -IsSiteCollectionAdmin $false 
    }
    catch {
        Write-Host "Unable to Remove $admin to  '$($site.url)' as Site Admin" -ForegroundColor Red
        Write-LogEntry -LogName:$Log -LogEntryText "Unable to Remove $admin to  '$($site.url)' as Site Admin"
    }
    Write-Host ""
    Write-LogEntry -LogName:$Log -LogEntryText ""
}

#Export the data to CSV
$output | Export-Csv $outputfile -NoTypeInformation -Append   
Write-Host " === === === === === Completed! === === === === === === == "
Write-Host "Collection output file $outputfile was saved" 
