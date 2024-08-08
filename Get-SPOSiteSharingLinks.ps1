<#   
.SYNOPSIS
    Get-SPOSiteSharingLinks.ps1 - Loops through all specified sites and exports all Sharing links for each site. 
      If the SPGroup Users field is empty, this mean the sharing link was never clicked on (redeemed). 

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
    Date: 8/8/2024
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
$t = 'admin' # < - Your Tenant Name Here
$admin = 'admin@contoso.com'  # <- Your Admin Account Here

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
$outputfile = "$env:TEMP\" + 'SPSite_SharingLinks_' + $date + '_' + "output.csv"
$log = "$env:TEMP\" + 'SPSite_SharingLinks_' + $date + '_' + "logfile.log"

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
        $AADGroups = Get-UnifiedGroup -Identity $siteprops.GroupId | Select-Object Guid, DisplayName, Alias, AccessType, WhenCreated
    }
 
    #Write-host "Information Barrier Mode:" $siteprops.InformationBarriersMode -ForegroundColor White
    Write-LogEntry -LogName:$Log -LogEntryText "Information Barrier Mode: $($siteprops.InformationBarriersMode)"
 
    #Write-host "Information Barrier Segment:" $siteprops.InformationSegment -ForegroundColor White
    Write-LogEntry -LogName:$Log -LogEntryText  "Information Barrier Segment: $($siteprops.InformationSegment)"

    #Get All Groups of a site collection
    $Groups = Get-SPOSiteGroup -Site $site.Url

    #Write-host "Total Number of Groups Found:" $Groups.Count -ForegroundColor White
    Write-LogEntry -LogName:$Log -LogEntryText "Total Number of Groups Found: $($Groups.Count)"

    ForEach ($Group in $Groups) {

        #If statement is to only collect sites that contain sharing links 
        if ($group.Title -like 'SharingLinks*') {
    
            #Write-Host "Group Title: $($Group.Title)" -ForegroundColor Yellow
            Write-LogEntry -LogName:$Log -LogEntryText "Group Title: '$($Group.Title)'"

            #Write-Host "Group Roles: $($Group.Roles) " -ForegroundColor Red
            Write-LogEntry -LogName:$Log -LogEntryText "Group Roles: '$($Group.Roles)'" 

            #Write-Host "Users in Group: '$($Group.Users)'" -ForegroundColor Cyan
            Write-LogEntry -LogName:$Log -LogEntryText "Users in Group: '$($Group.Users)'"

            Write-Host ""
            Write-LogEntry -LogName:$Log -LogEntryText ""

            # This script block iterates over each user in a group.
            # For each user, it attempts to retrieve the user's display name, primary SMTP address, and information barrier segments using the Get-Recipient cmdlet.
            # The retrieved information is then logged using a custom Write-LogEntry function.
            # If the script is unable to retrieve the information for a user, it prints an error message to the console and logs the error.  


            #Collecting Export Properties for CSV File
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
            
            $ExportItem  | Add-Member -MemberType NoteProperty -name "SPGroup Title" -value ($($Group.Title) -join ',')
            $ExportItem  | Add-Member -MemberType NoteProperty -name "SPGroup Users" -value ($($Group.Users) -join ',')

            $ExportItem  | Add-Member -MemberType NoteProperty -name "Group Owners" -value ($($groupowners.PrimarySmtpAddress) -join ',')
            $ExportItem  | Add-Member -MemberType NoteProperty -name "SPGroup Members" -value ($($groupmembers.PrimarySmtpAddress) -join ',')

            $ExportItem  | Add-Member -MemberType NoteProperty -name "Entra Group Displayname" -value ($($AADGroups.DisplayName))
            $ExportItem  | Add-Member -MemberType NoteProperty -name "Entra Group Alias" -value ($($AADGroups.Alias) -join ',')
            $ExportItem  | Add-Member -MemberType NoteProperty -name "Entra Group AccessType" -value ($($AADGroups.AccessType) -join ',')
            $ExportItem  | Add-Member -MemberType NoteProperty -name "Entra Group WhenCreated" -value ($($AADGroups.WhenCreated) -join ',')

            $output += $ExportItem

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
