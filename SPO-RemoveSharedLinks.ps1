# Connect to your SharePoint site
Connect-PnPOnline -Url https://m365x13453069.sharepoint.com/sites/Debra-Team-Ch1-Pri -interactive
 
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
