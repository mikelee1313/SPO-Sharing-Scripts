<#
.SYNOPSIS
This script connects to Exchange Online, retrieves unified groups based on access type filter, and exports their details to a CSV file.

.DESCRIPTION
The script performs the following actions:
1. Connects to Exchange Online.
2. Retrieves unified groups filtered by the specified AccessType and selects specific properties (Guid, Email, Alias, AccessType, WhenCreated,SharePointSiteUrl).
3. Creates a custom PowerShell object to store the group details.
4. Exports the group details to a CSV file in the TEMP directory with a filename that includes the current date.
5. Displays a completion message with the path to the output file.

.PARAMETER AccessType
Specifies which groups to retrieve based on their access type. Valid values are:
- "All" (default): Retrieves all groups regardless of access type
- "Public": Retrieves only public groups
- "Private": Retrieves only private groups

.PARAMETER $outputfile
The path to the output CSV file, which is stored in the TEMP directory with a filename that includes the current date.

.PARAMETER $groups
A collection of unified groups retrieved from Exchange Online.

.PARAMETER $ExportItem
A custom PowerShell object that stores the details of the unified groups.

.OUTPUT
A CSV file containing the details of unified groups filtered by access type.

.EXAMPLE
.\get-allgroups.ps1
This example runs the script with default AccessType "All" and exports all unified groups to a CSV file.

.EXAMPLE
.\get-allgroups.ps1 -AccessType "Public"
This example runs the script and exports only public unified groups to a CSV file.

.EXAMPLE
.\get-allgroups.ps1 -AccessType "Private"
This example runs the script and exports only private unified groups to a CSV file.

.NOTES
- Ensure you have the necessary permissions to connect to Exchange Online and retrieve unified groups.
- The script requires the Exchange Online PowerShell module to be installed and imported.
#>


# Define the output file path with the current date and access type filter
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss" # Current date and time for unique file naming
$outputfile = "$env:TEMP\" + 'All_Groups_' + $AccessType + '_' + $date + '_' + "output.csv"

#Modify the AccessType parameter to filter by your by access type preference. 
$AccessType = "all" #Example: Public, Private, or All.

# Connect to Exchange Online
Connect-ExchangeOnline 

Write-Host "Retrieving unified groups with AccessType filter: $AccessType" -ForegroundColor Green

# Retrieve unified groups based on AccessType parameter and select specific properties
if ($AccessType -eq "All") {
    $groups = Get-UnifiedGroup | Select-Object Guid, PrimarySmtpAddress, Alias, AccessType, WhenCreated, SharePointSiteUrl 
    Write-Host "Retrieved all unified groups (no filter applied)" -ForegroundColor Yellow
}
else {
    $groups = Get-UnifiedGroup | Where-Object { $_.AccessType -eq $AccessType } | Select-Object Guid, PrimarySmtpAddress, Alias, AccessType, WhenCreated, SharePointSiteUrl 
    Write-Host "Retrieved $($groups.Count) unified groups with AccessType: $AccessType" -ForegroundColor Yellow
} 

# Initialize an array to store the export items
$output = @()

# Iterate through each group and create a custom object with the group details
foreach ($group in $groups) {
    $ExportItem = New-Object PSObject
    $ExportItem | Add-Member -MemberType NoteProperty -Name "Alias" -Value $group.Alias
    $ExportItem | Add-Member -MemberType NoteProperty -Name "Email" -Value $group.PrimarySmtpAddress
    $ExportItem | Add-Member -MemberType NoteProperty -Name "GUID" -Value $group.Guid
    $ExportItem | Add-Member -MemberType NoteProperty -Name "AccessType" -Value $group.AccessType
    $ExportItem | Add-Member -MemberType NoteProperty -Name "SharePointSiteUrl" -Value $group.SharePointSiteUrl
    $ExportItem | Add-Member -MemberType NoteProperty -Name "WhenCreated" -Value $group.WhenCreated
    $output += $ExportItem
}

# Export the data to a CSV file
$output | Export-Csv $outputfile -NoTypeInformation   

# Display a completion message with the path to the output file
Write-Host " === === === === === Completed! === === === === === === == " -ForegroundColor Green
Write-Host "Exported $($output.Count) groups with AccessType '$AccessType' to:" -ForegroundColor Yellow
Write-Host "Collection output file $outputfile was saved" -ForegroundColor Cyan
