<#
.SYNOPSIS
This script removes sharing links for all files in a SharePoint library.

.DESCRIPTION
The script connects to a SharePoint site, specifies the library URL, retrieves all files in the library, 
and then removes the sharing links for each file. Finally, it disconnects from SharePoint.

.PARAMETER libraryUrl
The URL of the SharePoint library where the files are located.

.EXAMPLE
Remove-SharePointFileSharingLinks -libraryUrl "Shared Documents"
#>

# Connect to your SharePoint site
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/Team123 -interactive
 
# Specify the library URL
$libraryUrl = "Shared Documents"
 
# Get all files in the library
$files = Get-PnPListItem -List $libraryUrl
# Loop through each file and remove sharing links
foreach ($file in $files) {
    $fileUrl = $file.FieldValues["FileRef"]
    Remove-PnPFileSharingLink -FileUrl $fileUrl -Force
    Write-Host "Removed sharing links for file: $fileUrl"
}
# Disconnect from SharePoint
Disconnect-PnPOnline
