<#
.SYNOPSIS
    Scans SharePoint Online sites to identify all files and folders and their permissions.

.DESCRIPTION
    This script connects to SharePoint Online using provided tenant-level credentials and iterates through a list of 
    site URLs specified in an input file. It recursively scans document libraries and lists (excluding specified folders) 
    to locate all files and folders, and then details their permissions, including who has access (users/groups), 
    what roles they have, and whether permissions are unique or inherited.
    The script logs its operations and outputs the results to an Excel file using the ImportExcel module.

.PARAMETER None
    This script does not accept parameters via the command line. Configuration is done within the script.

.INPUTS
    A text file containing SharePoint site URLs to scan (path specified in $inputFilePath variable).

.OUTPUTS
    - An Excel file containing all found item permissions (path: $env:TEMP\All_Item_Permissions_[timestamp].xlsx)
    - A log file documenting the script's execution (path: $env:TEMP\All_Item_Permissions_[timestamp].txt)

.NOTES
    File Name      : Get-SPOItemPermissions.ps1
    Author         : Mike Lee
    Date Created   : 6/24/2025

    The script uses app-only authentication with a certificate thumbprint. Make sure the app has
    proper permissions in your tenant (Sites.FullControl.All is recommended).

    The script ignores several system folders and lists to improve performance and avoid errors.

    PREREQUISITES:
    - Install-Module ImportExcel -Scope CurrentUser
    - Install-Module PnP.PowerShell -Scope CurrentUser

.DISCLAIMER
Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
Microsoft further disclaims all implied warranties including, without limitation, 
any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, 
production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use of or inability 
to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

.EXAMPLE
    .\Get-SPOItemPermissions.ps1 # Assuming you rename the script
    Executes the script with the configured settings. Ensure you've updated the variables at the top
    of the script (appID, thumbprint, tenant, and inputFilePath) before running.
#>

# =================================================================================================
# USER CONFIGURATION - Update the variables in this section
# =================================================================================================

# --- Tenant and App Registration Details ---
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51"                # This is your Tenant ID

# --- Input File Path ---
$inputFilePath = 'C:\temp\SPOSiteList.txt' # Path to the input file containing site URLs

# --- Script Behavior Settings ---
$batchSize = 100  # How many items to process before writing to Excel
$maxItemsPerSheet = 5000 # Maximum items per sheet in Excel

# =================================================================================================
# END OF USER CONFIGURATION
# =================================================================================================

# Check for required modules
$requiredModules = @('ImportExcel', 'PnP.PowerShell')
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Module '$module' is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber | Out-Null
    }
    Import-Module $module -Force | Out-Null
}

# Script Parameters
Add-Type -AssemblyName System.Web
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\All_Item_Permissions_$startime.txt"
$outputFilePath = "$env:TEMP\All_Item_Permissions_$startime.xlsx"

# Initialize collections for batch processing
$global:currentBatch = @()
$global:totalItemsProcessed = 0
$global:currentSheetNumber = 1
$global:summaryData = @()
$global:excelFileInitialized = $false

# Setup logging
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
    
    # Also display important messages to console with color coding
    switch ($level) {
        "ERROR" { Write-Host $message -ForegroundColor Red }
        "WARNING" { Write-Host $message -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $message -ForegroundColor Green }
        default { 
            if ($level -eq "INFO" -and $message -match "Processing|Completed") {
                Write-Host $message -ForegroundColor Cyan
            }
        }
    }
}

# Handle SharePoint Online throttling with exponential backoff
function Invoke-WithRetry {
    param (
        [ScriptBlock]$ScriptBlock,
        [int]$MaxRetries = 5,
        [int]$InitialDelaySeconds = 5
    )
    
    $retryCount = 0
    $delay = $InitialDelaySeconds
    $success = $false
    $result = $null
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            $result = & $ScriptBlock
            [void]($success = $true)
        }
        catch {
            $exception = $_.Exception
            
            # Check if this is a throttling error (look for specific status codes or messages)
            [void]($isThrottlingError = $false)
            $retryAfterSeconds = $delay
            
            if ($null -ne $exception.Response) {
                # Check for Retry-After header
                $retryAfterHeader = $exception.Response.Headers['Retry-After']
                if ($retryAfterHeader) {
                    [void]($isThrottlingError = $true)
                    $retryAfterSeconds = [int]$retryAfterHeader
                    Write-Log "Received Retry-After header: $retryAfterSeconds seconds" "WARNING"
                }
                
                # Check for 429 (Too Many Requests) or 503 (Service Unavailable)
                $statusCode = [int]$exception.Response.StatusCode
                if ($statusCode -eq 429 -or $statusCode -eq 503) {
                    [void]($isThrottlingError = $true)
                    Write-Log "Detected throttling response (Status code: $statusCode)" "WARNING"
                }
            }
            
            # Also check for specific throttling error messages
            if ($exception.Message -match "throttl" -or 
                $exception.Message -match "too many requests" -or
                $exception.Message -match "temporarily unavailable") {
                [void]($isThrottlingError = $true)
                Write-Log "Detected throttling error in message: $($exception.Message)" "WARNING"
            }
            
            if ($isThrottlingError) {
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Log "Throttling detected. Retry attempt $retryCount of $MaxRetries. Waiting $retryAfterSeconds seconds..." "WARNING"
                    Write-Host "Throttling detected. Retry attempt $retryCount of $MaxRetries. Waiting $retryAfterSeconds seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfterSeconds
                    
                    # Implement exponential backoff if no Retry-After header was provided
                    if ($retryAfterSeconds -eq $delay) {
                        $delay = $delay * 2 # Exponential backoff
                    }
                }
                else {
                    Write-Log "Maximum retry attempts reached. Giving up on operation." "ERROR"
                    throw $_
                }
            }
            else {
                # Not a throttling error, rethrow
                $errorMessage = $_.Exception.Message
                $logLevel = "WARNING" # Default to WARNING for unexpected errors

                # Check for common, potentially less critical errors
                if ($errorMessage -match "File Not Found" -or $errorMessage -match "404" -or 
                    $errorMessage -match "Access denied" -or $errorMessage -match "403") {
                    $logLevel = "INFO" # Downgrade to INFO for these specific cases
                }
                Write-Log "General Error occurred During retrieval : $errorMessage" $logLevel
                throw $_
            }
        }
    }
    
    return $result
}

# Read site URLs from input file
function Read-SiteURLs {
    param (
        [string]$filePath
    )
    $urls = Get-Content -Path $filePath
    return $urls
}

# Connect to SharePoint Online
function Connect-SharePoint {
    param (
        [string]$siteURL
    )
    try {
        Invoke-WithRetry -ScriptBlock {
            Connect-PnPOnline -Url $siteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        }
        Write-Log "Connected to SharePoint Online at $siteURL"
        
        # Validate connection by trying to get the web
        try {
            $web = Get-PnPWeb -ErrorAction Stop
            Write-Log "Successfully validated connection to: $($web.Title) ($($web.Url))"
        }
        catch {
            Write-Log "Connection validation failed: $($_.Exception.Message)" "ERROR"
            return $false
        }
        
        return $true # Connection successful
    }
    catch {
        Write-Log "Failed to connect to SharePoint Online at $siteURL : $($_.Exception.Message)" "ERROR"
        return $false # Connection failed
    }
}

# List of folders to ignore
$ignoreFolders = @(
    "_catalogs",
    "appdata",
    "forms",
    "Form Templates",
    "Site Assets",
    "List Template Gallery",
    "Master Page Gallery",
    "Solution Gallery",
    "Style Library",
    "Composed Looks",
    "Converted Forms",
    "Web Part Gallery",
    "Theme Gallery",
    "TaxonomyHiddenList",
    "Events",
    "_cts",
    "_private",
    "_vti_pvt",
    "Reference 778a30bb4f074ae3bec315889ee34b88",
    "Sharing Links",
    "Social",
    "FavoriteLists-e0157a47-72e4-43c1-bfd0-ed9f7040e894",
    "User Information List",
    "Web Template Extensions",
    "SmartCache-8189C6B3-4081-4F62-9015-35FDB7FDF042",
    "SharePointHomeCacheList",
    "RecentLists-56BAEAB4-E7AD-4E59-B92B-9290D871F5C3",
    "PersonalCacheLibrary",
    "microsoft.ListSync.Endpoints",
    "Maintenance Log Library",
    "DO_NOT_DELETE_ENTERPRISE_USER_CONTAINER_ENUM_LIST_ee0de9c4-6398-408f-ac09-f0401edfb0bf",
    "appfiles",
    "(Reference, 778a30bb4f074ae3bec315889ee34b88)"
)

# Function to write batch data to Excel
function Write-BatchToExcel {
    param (
        [array]$Data,
        [string]$FilePath,
        [int]$SheetNumber
    )
    
    if ($Data.Count -eq 0) { return }
    
    try {
        $worksheetName = "File_Permissions_$SheetNumber"
        
        # Define Excel table style for better readability
        $excelParams = @{
            Path          = $FilePath
            WorksheetName = $worksheetName
            TableName     = "PermissionsTable$SheetNumber"
            TableStyle    = 'Medium6'
            AutoSize      = $true
            FreezeTopRow  = $true
            BoldTopRow    = $true
        }
        
        # Add conditional formatting for permission types
        $conditionalFormatting = @(
            New-ConditionalText -Text 'Unique' -BackgroundColor LightYellow -ConditionalTextColor Black
            New-ConditionalText -Text 'Inherited' -BackgroundColor LightGreen -ConditionalTextColor Black
        )
        
        # Export data to Excel. Create the file on first write, append on subsequent writes.
        if (-not $global:excelFileInitialized) {
            $Data | Export-Excel @excelParams -ConditionalText $conditionalFormatting
            [void]($global:excelFileInitialized = $true)
        }
        else {
            $Data | Export-Excel @excelParams -ConditionalText $conditionalFormatting -Append
        }
        
        Write-Log "Successfully wrote $($Data.Count) items to worksheet: $worksheetName" "SUCCESS"
    }
    catch {
        Write-Log "Failed to write batch to Excel: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Modified function to handle batch processing
function Add-ItemToBatch {
    param (
        [PSCustomObject]$Item
    )
    
    [void]($global:currentBatch += $Item)
    $global:totalItemsProcessed++
    
    # Check if we need to write the batch
    if ($global:currentBatch.Count -ge $batchSize) {
        Write-BatchToExcel -Data $global:currentBatch -FilePath $outputFilePath -SheetNumber $global:currentSheetNumber
        $global:currentBatch = @()
        
        # Check if we need a new sheet
        $itemsInCurrentSheet = ($global:totalItemsProcessed % $maxItemsPerSheet)
        if ($itemsInCurrentSheet -eq 0) {
            $global:currentSheetNumber++
        }
    }
    
    # Update progress every 10 items
    if ($global:totalItemsProcessed % 10 -eq 0) {
        Write-Host "Processed $global:totalItemsProcessed items..." -ForegroundColor Yellow
    }
}

# Process SharePoint Item (File or Folder)
function Get-SPItemPermission {
    param (
        $item,
        [string]$ItemSiteURL,
        [string]$ItemType, # "File" or "Folder"
        [string]$LibraryName
    )
    try {
        Write-Log "Getting permissions for $ItemType (ID: $($item.Id)) in list '$LibraryName'" "INFO"
        
        # The ParentList property is not loaded, causing an error. Use the passed-in parameter instead.
        # $libraryName = $item.ParentList.Title
        
        $itemName = ""
        $itemPath = ""

        # Access field values using indexer
        $itemName = $item["FileLeafRef"]
        $itemPath = $item["FileRef"]
        
        Write-Log "Processing item: $itemPath (Type: $ItemType)" "INFO"
        
        # Load role assignments
        Get-PnPProperty -ClientObject $item -Property RoleAssignments, HasUniqueRoleAssignments | Out-Null
       
        $permissionType = if ($item.HasUniqueRoleAssignments) { "Unique" } else { "Inherited" }

        # Creator and Created Date
        $creatorName = "Unknown"
        $creatorEmail = "Unknown"
        $createdDateStr = "Unknown"
        $creatorWithEmail = "Unknown"
        $createdDateTime = $null

        try {
            $authorField = $item["Author"]
            if ($null -ne $authorField) {
                if ($null -ne $authorField.LookupId) {
                    $creatorInfo = Get-PnPUser -Identity $authorField.LookupId -ErrorAction SilentlyContinue
                    if ($null -ne $creatorInfo) {
                        $creatorName = $creatorInfo.Title
                        $creatorEmail = $creatorInfo.Email
                        if ([string]::IsNullOrEmpty($creatorEmail)) {
                            $creatorWithEmail = $creatorName
                        }
                        else {
                            $creatorWithEmail = "$creatorName ($creatorEmail)"
                        }
                    }
                }
                elseif ($null -ne $authorField.LookupValue) {
                    $creatorName = $authorField.LookupValue
                    $creatorWithEmail = $creatorName
                }
            }
            
            $createdField = $item["Created"]
            if ($createdField) {
                $createdDateTime = $createdField
                $createdDateStr = $createdDateTime.ToString("yyyy-MM-dd HH:mm:ss")
            }
        }
        catch {
            Write-Log "Error retrieving creator/date for item $itemPath : $($_.Exception.Message)" "INFO"
        }

        # Collect all permissions for this item
        $allPrincipals = @()
        $allPrincipalsWithRoles = @()

        # Process RoleAssignments
        if ($item.RoleAssignments -and $item.RoleAssignments.Count -gt 0) {
            foreach ($RoleAssignment in $item.RoleAssignments) {
                Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member | Out-Null
                
                $principalDisplayName = $RoleAssignment.Member.Title 
                if ([string]::IsNullOrEmpty($principalDisplayName)) {
                    $principalDisplayName = $RoleAssignment.Member.LoginName 
                }
                $principalEmail = $RoleAssignment.Member.Email
                if ([string]::IsNullOrEmpty($principalEmail)) {
                    $principalEmail = "N/A"
                }
                
                # Filter out "Limited" permissions
                $assignedRoles = $RoleAssignment.RoleDefinitionBindings | 
                Where-Object { $_.Name -ne "Limited Access" } | 
                ForEach-Object { $_.Name }
                
                # Skip this principal if they only have Limited Access
                if ($assignedRoles.Count -eq 0) {
                    Write-Log "Skipping 'Limited Access' permission for $principalDisplayName on $itemPath" "INFO"
                    continue
                }
                
                $assignedRolesStr = $assignedRoles -join ", "
                
                [void]($allPrincipals += $principalDisplayName)
                
                if ($principalEmail -eq "N/A" -or [string]::IsNullOrEmpty($principalEmail)) {
                    [void]($allPrincipalsWithRoles += "${principalDisplayName}: ${assignedRolesStr}")
                }
                else {
                    [void]($allPrincipalsWithRoles += "$principalDisplayName ($principalEmail): ${assignedRolesStr}")
                }
            }
        }
        
        # Always create an entry, even if no specific permissions
        $permissionEntry = [PSCustomObject]@{
            SiteURL        = $ItemSiteURL
            ItemType       = $ItemType
            LibraryName    = $LibraryName
            ItemPath       = $itemPath 
            ItemName       = $itemName
            CreatedBy      = $creatorWithEmail
            CreatedDate    = $createdDateTime
            PermissionType = $permissionType
            UserCount      = $allPrincipals.Count
            Permissions    = if ($allPrincipalsWithRoles.Count -gt 0) { ($allPrincipalsWithRoles -join "`n") } else { "Inherited from parent" }
        }
        
        Add-ItemToBatch -Item $permissionEntry
        Write-Log "Added item to batch: $itemName" "INFO"
    }
    catch {
        $itemId = try { $item.Id } catch { "Unknown" }
        Write-Log "Failed to process $ItemType (ID: $itemId): $($_.Exception.Message)" "ERROR"
        Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    }
}

# Function to create summary worksheet
function Create-SummaryWorksheet {
    param (
        [string]$FilePath
    )
    
    try {
        # Create summary data
        $summary = [PSCustomObject]@{
            'Total Sites Processed'         = $global:summaryData.Count
            'Total Items Processed'         = $global:totalItemsProcessed
            'Items with Unique Permissions' = ($global:summaryData | Where-Object { $_.UniquePermissions -gt 0 } | Measure-Object -Property UniquePermissions -Sum).Sum
            'Processing Start Time'         = $script:startTime
            'Processing End Time'           = Get-Date
            'Processing Duration'           = (Get-Date) - $script:startTime
        }
        
        # Export summary to first worksheet
        $summary | Export-Excel -Path $FilePath -WorksheetName "Summary" -TableName "SummaryTable" -TableStyle 'Medium2' -AutoSize -MoveToStart
        
        # Add site-level summary
        if ($global:summaryData.Count -gt 0) {
            $global:summaryData | Export-Excel -Path $FilePath -WorksheetName "Site Summary" -TableName "SiteSummaryTable" -TableStyle 'Medium4' -AutoSize -FreezeTopRow -BoldTopRow
        }
        
        Write-Log "Summary worksheet created successfully" "SUCCESS"
    }
    catch {
        Write-Log "Failed to create summary worksheet: $($_.Exception.Message)" "ERROR"
    }
}

# Main script execution
$script:startTime = Get-Date
Write-Log "Script started at $($script:startTime)"
Write-Log "Output will be saved to: $outputFilePath"

$siteURLs = Read-SiteURLs -filePath $inputFilePath
Write-Log "Found $($siteURLs.Count) sites to process"

foreach ($siteURL in $siteURLs) {
    $siteStartTime = Get-Date
    Write-Log "Starting processing for site: $siteURL" "INFO"
    
    $siteItemCount = 0
    $siteUniquePermissionCount = 0
    
    if (Connect-SharePoint -siteURL $siteURL) {
        try {
            # Get all lists including document libraries
            $lists = Get-PnPList -Includes BaseType, Hidden, Title, ItemCount | Where-Object { 
                $_.Hidden -eq $false -and 
                $_.Title -notin $ignoreFolders -and
                ($_.BaseType -eq "DocumentLibrary" -or $_.BaseType -eq "GenericList")
            }
            
            if ($null -eq $lists -or $lists.Count -eq 0) {
                Write-Log "No lists retrieved or all lists were ignored for site $siteURL." "WARNING"
                
                # Debug: Show all lists for troubleshooting
                $allLists = Get-PnPList -Includes Title, Hidden, BaseType
                Write-Log "Debug - All lists in site: $($allLists | ForEach-Object { "$($_.Title) (Hidden: $($_.Hidden), BaseType: $($_.BaseType))" } | Out-String)" "INFO"
            }
            else {
                Write-Log "Found $($lists.Count) lists to process in site $siteURL"
                Write-Log "Lists to process: $($lists | ForEach-Object { $_.Title } | Join-String -Separator ', ')" "INFO"
                
                foreach ($list in $lists) { 
                    try {
                        $listName = $list.Title
                        Write-Log "Processing list/library: '$listName' on site: $siteURL"
                        
                        # Get item count first
                        $itemCount = $list.ItemCount
                        Write-Log "List '$listName' contains $itemCount items"
                        
                        if ($itemCount -eq 0) {
                            Write-Log "Skipping empty list: $listName"
                            continue
                        }
                        
                        # Get all items at once with required fields
                        try {
                            Write-Log "Retrieving all items from list '$listName'..."
                            
                            $items = @(Get-PnPListItem -List $list -PageSize 2000)
                            
                            if ($null -eq $items -or $items.Count -eq 0) {
                                Write-Log "No items retrieved from list '$listName'" "WARNING"
                                continue
                            }
                            
                            Write-Log "Retrieved $($items.Count) items from list '$listName'"
                            $itemsProcessedInList = 0
                            
                            foreach ($currentItem in $items) {
                                try {
                                    # Get field values
                                    $fsObjType = $currentItem["FSObjType"]
                                    $itemTypeStr = ""
                                    
                                    if ($fsObjType -eq 0) {
                                        $itemTypeStr = "File"
                                    }
                                    elseif ($fsObjType -eq 1) {
                                        $itemTypeStr = "Folder"
                                    }
                                    else {
                                        Write-Log "Skipping item with unknown FSObjType: $fsObjType" "INFO"
                                        continue
                                    }
                                    
                                    $currentItemPath = $currentItem["FileRef"]
                                    
                                    # Check if item should be ignored
                                    [void]($ignoreCurrentItem = $false)
                                    foreach ($ignoreFolderPattern in $ignoreFolders) {
                                        if ($currentItemPath -like "*/$ignoreFolderPattern/*" -or $currentItemPath -match "/$($ignoreFolderPattern)$") {
                                            [void]($ignoreCurrentItem = $true)
                                            break
                                        }
                                    }
                                    
                                    if ($ignoreCurrentItem) {
                                        Write-Log "Ignoring item: $currentItemPath" "INFO"
                                        continue
                                    }
                                    
                                    Get-SPItemPermission -item $currentItem -ItemSiteURL $siteURL -ItemType $itemTypeStr -LibraryName $listName
                                    $siteItemCount++
                                    $itemsProcessedInList++
                                    
                                    # Check for unique permissions
                                    try {
                                        Get-PnPProperty -ClientObject $currentItem -Property HasUniqueRoleAssignments | Out-Null
                                        if ($currentItem.HasUniqueRoleAssignments) {
                                            $siteUniquePermissionCount++
                                        }
                                    }
                                    catch {
                                        Write-Log "Could not check unique permissions for item: $currentItemPath" "INFO"
                                    }
                                }
                                catch {
                                    Write-Log "Error processing individual item: $($_.Exception.Message)" "WARNING"
                                }
                            }
                            
                            Write-Log "Completed processing list '$listName'. Items processed: $itemsProcessedInList"
                        }
                        catch {
                            Write-Log "Error retrieving items from list '$listName': $($_.Exception.Message)" "ERROR"
                        }
                    }
                    catch {
                        Write-Log "Failed to process list '$($list.Title)' on site '$siteURL'. Error: $($_.Exception.Message)" "ERROR"
                    }
                }
            }
        }
        catch {
            Write-Log "Failed to get lists for site $siteURL. Error: $($_.Exception.Message)" "ERROR"
        }
    }
    
    # Add site summary data
    $siteSummary = [PSCustomObject]@{
        SiteURL           = $siteURL
        ItemsProcessed    = $siteItemCount
        UniquePermissions = $siteUniquePermissionCount
        ProcessingTime    = ((Get-Date) - $siteStartTime).ToString()
    }
    [void]($global:summaryData += $siteSummary)
    
    Write-Log "Completed processing for $siteURL. Items: $siteItemCount, Unique Permissions: $siteUniquePermissionCount" "SUCCESS"
}

# Write any remaining items in the batch - THIS IS CRITICAL
Write-Log "Writing final batch of $($global:currentBatch.Count) items"
if ($global:currentBatch.Count -gt 0) {
    Write-BatchToExcel -Data $global:currentBatch -FilePath $outputFilePath -SheetNumber $global:currentSheetNumber
    Write-Log "Final batch written successfully"
}

# Create summary worksheet only if we have data
if ($global:totalItemsProcessed -gt 0 -or $global:summaryData.Count -gt 0) {
    Create-SummaryWorksheet -FilePath $outputFilePath
}
else {
    Write-Log "No items were processed. Check if the sites contain any files or if permissions allow access." "WARNING"
}

# Final summary
$totalTime = (Get-Date) - $script:startTime
Write-Log "Item permissions scan completed. Total items processed: $global:totalItemsProcessed" "SUCCESS"
Write-Log "Total processing time: $totalTime"
Write-Log "Results available in: $outputFilePath" "SUCCESS"

# Check if file exists before trying to open
if (Test-Path $outputFilePath) {
    Write-Log "Excel file created successfully at: $outputFilePath"
    # Open the Excel file
    try {
        Start-Process $outputFilePath
    }
    catch {
        Write-Log "Could not automatically open the Excel file. Please open manually: $outputFilePath" "INFO"
    }
}
else {
    Write-Log "ERROR: Excel file was not created. Check the log for errors." "ERROR"
}
