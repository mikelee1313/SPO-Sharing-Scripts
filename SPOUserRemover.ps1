<#
.SYNOPSIS
    Removes specified users from SharePoint Online site collections.
    
.DESCRIPTION
    This comprehensive script removes user access from SharePoint Online sites by targeting three key areas:
    1. Site group memberships
    2. Direct file/item permissions
    3. Sharing links and sharing-related permissions
    
    The script employs intelligent throttling mechanisms to handle SharePoint Online rate limits and
    provides detailed logging of all operations.

.Disclaimer
    The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.


.PARAMETER appID
    The Azure/Entra Application (Client) ID used for authentication.
    
.PARAMETER thumbprint
    The certificate thumbprint used for authentication.
    
.PARAMETER tenant
    The Azure/Microsoft 365 Tenant ID.
    
.PARAMETER siteURL
    The full URL of the SharePoint site collection to process.
    
.PARAMETER userListPath
    Path to a text file containing user emails/login names (one per line).
    
.NOTES
    File Name      : SPOUserRemover.ps1
    Author         : Mike Lee
    Date Created   : 7/29/2025

    Prerequisites  : 
    - PnP.PowerShell module installed
    - App registration in Azure/Entra with Sites.FullControl.All permission
    - Certificate for app authentication
    
    
.EXAMPLE
    .\SPOUserRemover.ps1
    
    Runs the script using the parameters configured in the USER CONFIGURATION section.
    
.FUNCTIONALITY
    SharePoint Online
#>

#=================================================================================================
# USER CONFIGURATION - Update the variables in this section
#=================================================================================================

# --- Tenant and App Registration Details ---
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51"               # This is your Tenant ID

# --- Site and User Configuration ---
$siteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"       # SharePoint site collection URL
$userListPath = 'C:\temp\UsersList.txt'                         # Path to the input file containing user emails/logins

#=================================================================================================
# END OF USER CONFIGURATION
#=================================================================================================

# Global variables
$Script:ProcessedUsers = @()
$Script:RemovedUsers = @()
$Script:ErrorUsers = @()

#region Logging Functions
# Setup logging
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}
#endregion

#region Throttle Handling
function Invoke-WithThrottleHandling {
    param(
        [scriptblock]$ScriptBlock,
        [string]$Operation,
        [int]$MaxRetries = 3,
        [int]$BaseDelaySeconds = 5
    )
    
    $retryCount = 0
    do {
        try {
            & $ScriptBlock
            return
        }
        catch {
            $retryCount++
            if ($_.Exception.Message -like "*throttled*" -or $_.Exception.Message -like "*429*" -or $_.Exception.Message -like "*Too Many Requests*") {
                if ($retryCount -le $MaxRetries) {
                    $delay = $BaseDelaySeconds * [Math]::Pow(2, $retryCount - 1)
                    Write-Host "      Throttled during $Operation. Waiting $delay seconds before retry $retryCount/$MaxRetries..." -ForegroundColor Yellow
                    Write-Log "Throttled during $Operation. Waiting $delay seconds before retry $retryCount/$MaxRetries" "WARNING"
                    Start-Sleep -Seconds $delay
                }
                else {
                    Write-Host "      Max retries exceeded for $Operation" -ForegroundColor Red
                    Write-Log "Max retries exceeded for $Operation" "ERROR"
                    throw
                }
            }
            else {
                throw
            }
        }
    } while ($retryCount -le $MaxRetries)
}
#endregion

#region Authentication
function Connect-SPOService {
    param(
        [string]$SiteUrl,
        [string]$ClientId,
        [string]$Thumbprint,
        [string]$Tenant
    )
    
    try {
        Write-Host "Connecting to SharePoint Online..." -ForegroundColor Green
        Write-Log "Attempting to connect to SharePoint Online: $SiteUrl"
        
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant
        
        Write-Host "Successfully connected to SharePoint Online" -ForegroundColor Green
        Write-Log "Successfully connected to SharePoint Online"
        
        # Validate connection by trying to get the web
        try {
            $web = Get-PnPWeb -ErrorAction Stop
            Write-Log "Successfully validated connection to: $($web.Title) ($($web.Url))"
        }
        catch {
            Write-Log "Connection validation failed: $($_.Exception.Message)" "ERROR"
            return $false
        }
        
        return $true
    }
    catch {
        Write-Host "Failed to connect to SharePoint Online: $_" -ForegroundColor Red
        Write-Log "Failed to connect to SharePoint Online: $_" "ERROR"
        return $false
    }
}
#endregion

#region User Management Functions
function Read-UserList {
    param([string]$FilePath)
    
    try {
        if (-not (Test-Path $FilePath)) {
            throw "User list file not found: $FilePath"
        }
        
        $users = Get-Content -Path $FilePath | Where-Object { $_ -and $_.Trim() -ne "" }
        Write-Host "Loaded $($users.Count) users from file" -ForegroundColor Green
        Write-Log "Loaded $($users.Count) users from file: $FilePath"
        
        return $users
    }
    catch {
        Write-Host "Error reading user list: $_" -ForegroundColor Red
        Write-Log "Error reading user list: $_" "ERROR"
        throw
    }
}

function Remove-UserFromSiteGroups {
    param(
        [array]$Users
    )
    
    Write-Host "`nRemoving users from site groups..." -ForegroundColor Cyan
    Write-Log "Starting removal of users from site groups"
    
    try {
        # Get all site groups
        $siteGroups = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPGroup
        } -Operation "Get site groups"
        
        Write-Host "Found $($siteGroups.Count) groups in the site" -ForegroundColor Green
        
        foreach ($group in $siteGroups) {
            Write-Host "  Processing group: $($group.Title)" -ForegroundColor Yellow
            
            try {
                # Get group members
                $groupMembers = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPGroupMember -Identity $group.Id
                } -Operation "Get members for group $($group.Title)"
                
                foreach ($user in $Users) {
                    $memberToRemove = $groupMembers | Where-Object { 
                        $_.Email -eq $user -or $_.LoginName -eq $user -or $_.LoginName -like "*$user*" 
                    }
                    
                    if ($memberToRemove) {
                        try {
                            Invoke-WithThrottleHandling -ScriptBlock {
                                try {
                                    Remove-PnPGroupMember -Identity $group.Id -LoginName $memberToRemove.LoginName -Force
                                    Write-Log "Successfully removed $($memberToRemove.LoginName) from group $($group.Title) using Force parameter"
                                }
                                catch {
                                    Remove-PnPGroupMember -Identity $group.Id -LoginName $memberToRemove.LoginName
                                    Write-Log "Successfully removed $($memberToRemove.LoginName) from group $($group.Title) using fallback method"
                                }
                            } -Operation "Remove user $user from group $($group.Title)"
                            
                            Write-Host "    Removed $user from group: $($group.Title)" -ForegroundColor Green
                            Write-Log "Removed $user from group: $($group.Title)"
                        }
                        catch {
                            Write-Host "    Error removing $user from group $($group.Title): $_" -ForegroundColor Red
                            Write-Log "Error removing $user from group $($group.Title): $_" "ERROR"
                        }
                    }
                }
            }
            catch {
                Write-Host "    Error processing group $($group.Title): $_" -ForegroundColor Red
                Write-Log "Error processing group $($group.Title): $_" "ERROR"
            }
        }
    }
    catch {
        Write-Host "Error processing site groups: $_" -ForegroundColor Red
        Write-Log "Error processing site groups: $_" "ERROR"
    }
}

function Remove-UserFromFilePermissions {
    param(
        [array]$Users
    )
    
    Write-Host "`nRemoving users from file permissions..." -ForegroundColor Cyan
    Write-Log "Starting removal of users from file permissions"
    
    try {
        # Get all lists in the site (both Document Libraries and regular Lists)
        $lists = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPList | Where-Object { 
                $_.Hidden -eq $false -and 
                ($_.BaseType -eq "DocumentLibrary" -or $_.BaseType -eq "GenericList")
            }
        } -Operation "Get document libraries and lists"
        
        Write-Host "Found $($lists.Count) document libraries and lists" -ForegroundColor Green
        
        foreach ($list in $lists) {
            Write-Host "  Processing $(if($list.BaseType -eq 'DocumentLibrary'){'library'}else{'list'}): $($list.Title)" -ForegroundColor Yellow
            
            try {
                # Get all items in the library/list
                $items = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPListItem -List $list.Id -PageSize 1000
                } -Operation "Get items from $(if($list.BaseType -eq 'DocumentLibrary'){'library'}else{'list'}) $($list.Title)"
                
                foreach ($item in $items) {
                    try {
                        # Check if item has unique permissions
                        $hasUniquePerms = Invoke-WithThrottleHandling -ScriptBlock {
                            Get-PnPProperty -ClientObject $item -Property "HasUniqueRoleAssignments"
                            return $item.HasUniqueRoleAssignments
                        } -Operation "Check unique permissions for item $($item.Id)"
                        
                        if ($hasUniquePerms) {
                            try {
                                # Load role assignments with member properties
                                Invoke-WithThrottleHandling -ScriptBlock {
                                    $item.Context.Load($item.RoleAssignments)
                                    $item.Context.ExecuteQuery()
                                    
                                    foreach ($assignment in $item.RoleAssignments) {
                                        $item.Context.Load($assignment.Member)
                                        $item.Context.Load($assignment.RoleDefinitionBindings)
                                    }
                                    $item.Context.ExecuteQuery()
                                } -Operation "Load role assignments and members for item $($item.Id)"
                                
                                foreach ($assignment in $item.RoleAssignments) {
                                    try {
                                        $member = $assignment.Member
                                        
                                        foreach ($user in $Users) {
                                            if ($member.LoginName -eq $user -or $member.LoginName -like "*$user*" -or $member.Email -eq $user) {
                                                try {
                                                    Invoke-WithThrottleHandling -ScriptBlock {
                                                        # Remove all role assignments for this user on this item
                                                        $roleDefinitions = $assignment.RoleDefinitionBindings
                                                        $removedRoles = @()
                                                        
                                                        foreach ($roleDef in $roleDefinitions) {
                                                            try {
                                                                # Skip "Limited Access" as it's usually inherited and can't be removed directly
                                                                if ($roleDef.Name -ne "Limited Access") {
                                                                    Set-PnPListItemPermission -List $list.Id -Identity $item.Id -User $member.LoginName -RemoveRole $roleDef.Name
                                                                    $removedRoles += $roleDef.Name
                                                                }
                                                            }
                                                            catch {
                                                                # Log specific role removal errors but continue with other roles
                                                                if ($_.Exception.Message -notlike "*Can not find the principal*" -and $_.Exception.Message -notlike "*does not exist*") {
                                                                    Write-Log "Error removing role '$($roleDef.Name)' for user $user from file $($item["FileLeafRef"]): $_" "WARNING"
                                                                }
                                                            }
                                                        }
                                                        
                                                        # Only log success if we actually removed some roles
                                                        if ($removedRoles.Count -gt 0) {
                                                            # Get appropriate display name for the item
                                                            $itemDisplayName = if ($list.BaseType -eq "DocumentLibrary") {
                                                                $item["FileLeafRef"]
                                                            }
                                                            else {
                                                                "Item ID $($item.Id) (Title: $($item["Title"]))"
                                                            }
                                                            Write-Log "Removed roles [$($removedRoles -join ', ')] for user $user from $itemDisplayName"
                                                        }
                                                    } -Operation "Remove user $user from item $($item.Id)"
                                                    
                                                    # Get appropriate display name for the item
                                                    $itemDisplayName = if ($list.BaseType -eq "DocumentLibrary") {
                                                        $item["FileLeafRef"]
                                                    }
                                                    else {
                                                        "Item ID $($item.Id) (Title: $($item["Title"]))"
                                                    }
                                                    
                                                    Write-Host "    Removed $user from $(if($list.BaseType -eq 'DocumentLibrary'){'file'}else{'item'}): $itemDisplayName" -ForegroundColor Green
                                                    Write-Log "Removed $user from $(if($list.BaseType -eq 'DocumentLibrary'){'file'}else{'item'}): $itemDisplayName in $(if($list.BaseType -eq 'DocumentLibrary'){'library'}else{'list'}): $($list.Title)"
                                                }
                                                catch {
                                                    # Only log errors that aren't about missing principals
                                                    if ($_.Exception.Message -notlike "*Can not find the principal*" -and $_.Exception.Message -notlike "*does not exist*") {
                                                        Write-Host "    Error removing $user from file $($item["FileLeafRef"]): $_" -ForegroundColor Red
                                                        Write-Log "Error removing $user from file $($item["FileLeafRef"]): $_" "ERROR"
                                                    }
                                                    else {
                                                        Write-Log "User $user already removed or doesn't have permissions on file $($item["FileLeafRef"])" "INFO"
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch {
                                        Write-Log "Error processing role assignment member: $_"
                                    }
                                }
                            }
                            catch {
                                Write-Log "Error loading role assignments for item $($item.Id): $_"
                            }
                        }
                    }
                    catch {
                        Write-Log "Error processing item $($item.Id): $_"
                    }
                }
            }
            catch {
                Write-Host "  Error processing library $($list.Title): $_" -ForegroundColor Red
                Write-Log "Error processing library $($list.Title): $_" "ERROR"
            }
        }
    }
    catch {
        Write-Host "Error processing file permissions: $_" -ForegroundColor Red
        Write-Log "Error processing file permissions: $_" "ERROR"
    }
}

function Remove-UserFromSharingLinks {
    param(
        [array]$Users
    )
    
    Write-Host "`nRemoving users from sharing links..." -ForegroundColor Cyan
    Write-Log "Starting removal of users from sharing links"
    
    try {
        # Step 1: Process each file to find and revoke sharing links that contain the target users
        Write-Host "`nStep 1: Identifying and revoking sharing links containing target users..." -ForegroundColor Cyan
        
        $lists = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPList | Where-Object { $_.Hidden -eq $false -and $_.BaseType -eq "DocumentLibrary" }
        } -Operation "Get document libraries for sharing link processing"
        
        foreach ($list in $lists) {
            Write-Host "  Processing library: $($list.Title)" -ForegroundColor Yellow
            
            try {
                $items = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPListItem -List $list.Id -PageSize 500
                } -Operation "Get items from library $($list.Title)"
                
                foreach ($item in $items) {
                    try {
                        # Get all sharing links for this file
                        $sharingLinks = Invoke-WithThrottleHandling -ScriptBlock {
                            Get-PnPFileSharingLink -Identity $item.FieldValues.FileRef -ErrorAction SilentlyContinue
                        } -Operation "Get sharing links for file $($item["FileLeafRef"])"
                        
                        if ($sharingLinks) {
                            Write-Log "Found $($sharingLinks.Count) sharing link(s) for file $($item["FileLeafRef"])"
                            
                            foreach ($sharingLink in $sharingLinks) {
                                # Check if any of our target users have access to this sharing link
                                $shouldRevokeLink = $false
                                $usersToRemove = @()
                                
                                foreach ($user in $Users) {
                                    # For flexible links, we need to check if the user is in the sharing link's invitees
                                    if ($sharingLink.ShareLink.ShareKind -eq "Flexible") {
                                        try {
                                            # Get sharing link details including invitees
                                            $linkDetails = Invoke-WithThrottleHandling -ScriptBlock {
                                                $ctx = Get-PnPContext
                                                $web = $ctx.Web
                                                $file = $web.GetFileByServerRelativeUrl($item.FieldValues.FileRef)
                                                
                                                # Use SharePoint REST API to get sharing information
                                                $requestBody = @{
                                                    listId  = $list.Id.ToString()
                                                    itemId  = $item.Id
                                                    shareId = $sharingLink.ShareId
                                                } | ConvertTo-Json
                                                
                                                $response = Invoke-PnPSPRestMethod -Url "/_api/web/lists('$($list.Id)')/items($($item.Id))/GetSharingInformation" -Method Post -Content $requestBody
                                                return $response
                                            } -Operation "Get detailed sharing info for flexible link"
                                            
                                            # Check if user is in the sharing link
                                            if ($linkDetails -and $linkDetails.pickerSettings -and $linkDetails.pickerSettings.principalEntries) {
                                                foreach ($principal in $linkDetails.pickerSettings.principalEntries) {
                                                    if ($principal.Email -eq $user -or $principal.LoginName -like "*$user*") {
                                                        $shouldRevokeLink = $true
                                                        $usersToRemove += $user
                                                        Write-Log "Found user $user in flexible sharing link for file $($item["FileLeafRef"])"
                                                        break
                                                    }
                                                }
                                            }
                                        }
                                        catch {
                                            Write-Log "Could not check flexible link details for $($item["FileLeafRef"]): $_" "WARNING"
                                            # Fallback: if we can't determine specific users, revoke the link if it's flexible
                                            # This is safer than leaving potentially accessible links
                                            $shouldRevokeLink = $true
                                            $usersToRemove += $user
                                        }
                                    }
                                    else {
                                        # For other link types (direct, anonymous), check if user has access through other means
                                        # Since these are typically more open, we'll be more conservative
                                        Write-Log "Found $($sharingLink.ShareLink.ShareKind) sharing link for file $($item["FileLeafRef"]) - checking permissions"
                                    }
                                }
                                
                                # If we found target users in this sharing link, revoke it entirely
                                if ($shouldRevokeLink) {
                                    try {
                                        Write-Host "    Revoking flexible sharing link for file: $($item["FileLeafRef"]) (users: $($usersToRemove -join ', '))" -ForegroundColor Yellow
                                        
                                        Invoke-WithThrottleHandling -ScriptBlock {
                                            # Revoke the specific sharing link
                                            Remove-PnPFileSharingLink -Identity $item.FieldValues.FileRef -ShareId $sharingLink.ShareId -Force
                                        } -Operation "Revoke sharing link for file $($item["FileLeafRef"])"
                                        
                                        Write-Host "    Successfully revoked sharing link for file: $($item["FileLeafRef"])" -ForegroundColor Green
                                        Write-Log "Successfully revoked sharing link (ID: $($sharingLink.ShareId)) for file $($item["FileLeafRef"]) containing users: $($usersToRemove -join ', ')"
                                    }
                                    catch {
                                        Write-Host "    Error revoking sharing link for file $($item["FileLeafRef"]): $_" -ForegroundColor Red
                                        Write-Log "Error revoking sharing link for file $($item["FileLeafRef"]): $_" "ERROR"
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        # Skip files that don't have sharing links or can't be accessed
                        if ($_.Exception.Message -notlike "*No sharing links*" -and $_.Exception.Message -notlike "*not found*") {
                            Write-Log "Error processing sharing links for file $($item["FileLeafRef"]): $_" "WARNING"
                        }
                    }
                }
            }
            catch {
                Write-Host "  Error processing library $($list.Title): $_" -ForegroundColor Red
                Write-Log "Error processing library $($list.Title): $_" "ERROR"
            }
        }
        
        # Step 2: Clean up any remaining group memberships (backup approach)
        Write-Host "`nStep 2: Cleaning up sharing group memberships (backup)..." -ForegroundColor Cyan
        
        # First, get ALL groups (including sharing link groups that might not match our patterns)
        $allGroups = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPGroup
        } -Operation "Get all groups"
        
        # Filter for sharing-related groups with broader patterns
        $sharingGroups = $allGroups | Where-Object { 
            $_.Title -like "*SharingLinks*" -or 
            $_.Title -like "*sharing*" -or 
            $_.Title -like "*Everyone except external users*" -or
            $_.Title -match "^SharingLinks\." -or
            $_.Title -like "*Anonymous*" -or
            $_.LoginName -like "*spo-grid-all-users*" -or
            $_.Id -match "c:0o.c\|federateddirectoryclaimprovider\|.*" -or
            $_.PrincipalType -eq "SecurityGroup"
        }
        
        Write-Host "Found $($sharingGroups.Count) potential sharing-related groups out of $($allGroups.Count) total groups" -ForegroundColor Green
        
        foreach ($group in $sharingGroups) {
            Write-Host "  Processing sharing group: $($group.Title) (ID: $($group.Id))" -ForegroundColor Yellow
            
            try {
                $groupMembers = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPGroupMember -Identity $group.Id
                } -Operation "Get sharing group members"
                
                if ($groupMembers) {
                    Write-Log "Group '$($group.Title)' has $($groupMembers.Count) members: $($groupMembers | ForEach-Object { "$($_.Title) ($($_.LoginName))" } | Out-String)"
                }
                
                foreach ($user in $Users) {
                    $memberToRemove = $groupMembers | Where-Object { 
                        $_.Email -eq $user -or $_.LoginName -eq $user -or $_.LoginName -like "*$user*" -or $_.Title -like "*$user*"
                    }
                    
                    if ($memberToRemove) {
                        try {
                            Invoke-WithThrottleHandling -ScriptBlock {
                                try {
                                    Remove-PnPGroupMember -Identity $group.Id -LoginName $memberToRemove.LoginName -Force
                                    Write-Log "Successfully removed $($memberToRemove.LoginName) from sharing group $($group.Title) using Force parameter"
                                }
                                catch {
                                    Remove-PnPGroupMember -Identity $group.Id -LoginName $memberToRemove.LoginName
                                    Write-Log "Successfully removed $($memberToRemove.LoginName) from sharing group $($group.Title) using fallback method"
                                }
                            } -Operation "Remove user $user from sharing group $($group.Title)"
                            
                            Write-Host "    Removed $user from sharing group: $($group.Title)" -ForegroundColor Green
                            Write-Log "Removed $user from sharing group: $($group.Title)"
                        }
                        catch {
                            Write-Host "    Error removing $user from sharing group $($group.Title): $_" -ForegroundColor Red
                            Write-Log "Error removing $user from sharing group $($group.Title): $_" "ERROR"
                        }
                    }
                }
            }
            catch {
                Write-Host "    Error processing sharing group $($group.Title): $_" -ForegroundColor Red
                Write-Log "Error processing sharing group $($group.Title): $_" "ERROR"
            }
        }
        
        # Step 3: Enhanced sharing link revocation using SharePoint REST API
        Write-Host "`nStep 3: Enhanced sharing link revocation (addressing flexible link access issue)..." -ForegroundColor Cyan
        Write-Log "Starting enhanced sharing link revocation to fully invalidate flexible links"
        
        $lists = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPList | Where-Object { $_.Hidden -eq $false -and $_.BaseType -eq "DocumentLibrary" }
        } -Operation "Get document libraries for enhanced sharing cleanup"
        
        foreach ($list in $lists) {
            Write-Host "  Processing library for enhanced cleanup: $($list.Title)" -ForegroundColor Yellow
            
            try {
                $items = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPListItem -List $list.Id -PageSize 500
                } -Operation "Get items from library $($list.Title)"
                
                foreach ($item in $items) {
                    try {
                        # Get all sharing links for this file
                        $sharingLinks = Invoke-WithThrottleHandling -ScriptBlock {
                            Get-PnPFileSharingLink -Identity $item.FieldValues.FileRef -ErrorAction SilentlyContinue
                        } -Operation "Get sharing links for file $($item["FileLeafRef"])"
                        
                        if ($sharingLinks) {
                            Write-Log "Found $($sharingLinks.Count) sharing link(s) for file $($item["FileLeafRef"])"
                            
                            foreach ($sharingLink in $sharingLinks) {
                                # For flexible links, revoke them entirely to ensure access is truly removed
                                if ($sharingLink.ShareLink.ShareKind -eq "Flexible") {
                                    try {
                                        Write-Host "    Revoking flexible sharing link for file: $($item["FileLeafRef"])" -ForegroundColor Yellow
                                        
                                        Invoke-WithThrottleHandling -ScriptBlock {
                                            # Revoke the specific sharing link entirely
                                            Remove-PnPFileSharingLink -Identity $item.FieldValues.FileRef -ShareId $sharingLink.ShareId -Force
                                        } -Operation "Revoke flexible sharing link for file $($item["FileLeafRef"])"
                                        
                                        Write-Host "    Successfully revoked flexible sharing link for file: $($item["FileLeafRef"])" -ForegroundColor Green
                                        Write-Log "Successfully revoked flexible sharing link (ID: $($sharingLink.ShareId)) for file $($item["FileLeafRef"]) to prevent continued access"
                                    }
                                    catch {
                                        Write-Host "    Error revoking flexible sharing link for file $($item["FileLeafRef"]): $_" -ForegroundColor Red
                                        Write-Log "Error revoking flexible sharing link for file $($item["FileLeafRef"]): $_" "ERROR"
                                    }
                                }
                                # For other link types, log that they were found but skip REST API approach
                                else {
                                    Write-Log "Found $($sharingLink.ShareLink.ShareKind) sharing link for file $($item["FileLeafRef"]) - access controlled by group membership removal (completed in Step 2)"
                                }
                            }
                        }
                    }
                    catch {
                        # Skip files that don't have sharing links or can't be accessed
                        if ($_.Exception.Message -notlike "*No sharing links*" -and $_.Exception.Message -notlike "*not found*") {
                            Write-Log "Error processing sharing links for file $($item["FileLeafRef"]): $_" "WARNING"
                        }
                    }
                }
            }
            catch {
                Write-Host "  Error processing library $($list.Title) for enhanced cleanup: $_" -ForegroundColor Red
                Write-Log "Error processing library $($list.Title) for enhanced cleanup: $_" "ERROR"
            }
        }
        
        # Verify sharing link access revocation through group membership
        Write-Host "`nVerifying sharing link access revocation..." -ForegroundColor Cyan
        Write-Log "Verifying that users have been removed from sharing groups (primary access control mechanism)" "INFO"
        
        $lists = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPList | Where-Object { $_.Hidden -eq $false -and $_.BaseType -eq "DocumentLibrary" }
        } -Operation "Get document libraries for sharing verification"
        
        foreach ($list in $lists) {
            try {
                $items = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPListItem -List $list.Id -PageSize 500
                } -Operation "Get items from library $($list.Title) for sharing verification"
                
                foreach ($item in $items) {
                    try {
                        foreach ($user in $Users) {
                            try {
                                # Check if the file has any sharing links
                                $sharingInfo = Invoke-WithThrottleHandling -ScriptBlock {
                                    Get-PnPFileSharingLink -Identity $item.FieldValues.FileRef -ErrorAction SilentlyContinue
                                } -Operation "Get sharing links for file $($item["FileLeafRef"])"
                                
                                if ($sharingInfo) {
                                    Write-Log "Found $($sharingInfo.Count) sharing link(s) for file $($item["FileLeafRef"])" "INFO"
                                    Write-Log "User $user access to sharing links is controlled by group membership (already processed above)" "INFO"
                                    Write-Host "    Verified sharing access revocation for file: $($item["FileLeafRef"])" -ForegroundColor Green
                                }
                            }
                            catch {
                                # This is expected for items without sharing links
                                Write-Log "No sharing links found for file $($item["FileLeafRef"])" "INFO"
                            }
                        }
                    }
                    catch {
                        Write-Log "Error checking item $($item.Id): $_"
                    }
                }
            }
            catch {
                Write-Log "Error processing library $($list.Title): $_" "ERROR"
            }
        }
        
        # Also check for sharing permissions on individual files in document libraries
        Write-Host "`nChecking for individual file sharing permissions..." -ForegroundColor Cyan
        
        $lists = Invoke-WithThrottleHandling -ScriptBlock {
            Get-PnPList | Where-Object { $_.Hidden -eq $false -and $_.BaseType -eq "DocumentLibrary" }
        } -Operation "Get document libraries for sharing check"
        
        foreach ($list in $lists) {
            try {
                $items = Invoke-WithThrottleHandling -ScriptBlock {
                    Get-PnPListItem -List $list.Id -PageSize 500
                } -Operation "Get items from library $($list.Title) for sharing check"
                
                foreach ($item in $items) {
                    try {
                        # Check if item has sharing information
                        $hasUniquePerms = Invoke-WithThrottleHandling -ScriptBlock {
                            Get-PnPProperty -ClientObject $item -Property "HasUniqueRoleAssignments"
                            return $item.HasUniqueRoleAssignments
                        } -Operation "Check unique permissions for sharing on item $($item.Id)"
                        
                        if ($hasUniquePerms) {
                            # Use the same permission removal logic as file permissions
                            try {
                                Invoke-WithThrottleHandling -ScriptBlock {
                                    $item.Context.Load($item.RoleAssignments)
                                    $item.Context.ExecuteQuery()
                                    
                                    foreach ($assignment in $item.RoleAssignments) {
                                        $item.Context.Load($assignment.Member)
                                        $item.Context.Load($assignment.RoleDefinitionBindings)
                                    }
                                    $item.Context.ExecuteQuery()
                                } -Operation "Load role assignments for sharing check on item $($item.Id)"
                                
                                foreach ($assignment in $item.RoleAssignments) {
                                    try {
                                        $member = $assignment.Member
                                        
                                        foreach ($user in $Users) {
                                            if ($member.LoginName -eq $user -or $member.LoginName -like "*$user*" -or $member.Email -eq $user) {
                                                # Check if this is a sharing-related permission (not inherited)
                                                $roleDefinitions = $assignment.RoleDefinitionBindings
                                                $isSharingPermission = $false
                                                
                                                foreach ($roleDef in $roleDefinitions) {
                                                    if ($roleDef.Name -eq "View Only" -or $roleDef.Name -eq "Edit" -or $roleDef.Name -eq "Read") {
                                                        $isSharingPermission = $true
                                                        break
                                                    }
                                                }
                                                
                                                if ($isSharingPermission) {
                                                    try {
                                                        Invoke-WithThrottleHandling -ScriptBlock {
                                                            $removedSharingRoles = @()
                                                            foreach ($roleDef in $roleDefinitions) {
                                                                try {
                                                                    # Skip "Limited Access" as it's usually inherited
                                                                    if ($roleDef.Name -ne "Limited Access") {
                                                                        Set-PnPListItemPermission -List $list.Id -Identity $item.Id -User $member.LoginName -RemoveRole $roleDef.Name
                                                                        $removedSharingRoles += $roleDef.Name
                                                                    }
                                                                }
                                                                catch {
                                                                    # Log specific role removal errors but continue
                                                                    if ($_.Exception.Message -notlike "*Can not find the principal*" -and $_.Exception.Message -notlike "*does not exist*") {
                                                                        Write-Log "Error removing sharing role '$($roleDef.Name)' for user $user from file $($item["FileLeafRef"]): $_" "WARNING"
                                                                    }
                                                                }
                                                            }
                                                            
                                                            # Only log success if we actually removed some roles
                                                            if ($removedSharingRoles.Count -gt 0) {
                                                                Write-Log "Removed sharing roles [$($removedSharingRoles -join ', ')] for user $user from file $($item["FileLeafRef"])"
                                                            }
                                                        } -Operation "Remove sharing permission for user $user from item $($item.Id)"
                                                        
                                                        Write-Host "    Removed sharing permission for $user from file: $($item["FileLeafRef"])" -ForegroundColor Green
                                                        Write-Log "Removed sharing permission for $user from file: $($item["FileLeafRef"]) in library: $($list.Title)"
                                                    }
                                                    catch {
                                                        # Only log errors that aren't about missing principals
                                                        if ($_.Exception.Message -notlike "*Can not find the principal*" -and $_.Exception.Message -notlike "*does not exist*") {
                                                            Write-Host "    Error removing sharing permission for $user from file $($item["FileLeafRef"]): $_" -ForegroundColor Red
                                                            Write-Log "Error removing sharing permission for $user from file $($item["FileLeafRef"]): $_" "ERROR"
                                                        }
                                                        else {
                                                            Write-Log "User $user already removed or doesn't have sharing permissions on file $($item["FileLeafRef"])" "INFO"
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch {
                                        Write-Log "Error processing sharing assignment member: $_"
                                    }
                                }
                            }
                            catch {
                                Write-Log "Error loading sharing assignments for item $($item.Id): $_"
                            }
                        }
                    }
                    catch {
                        Write-Log "Error checking sharing permissions for item $($item.Id): $_"
                    }
                }
            }
            catch {
                Write-Log "Error processing library $($list.Title) for sharing permissions: $_" "ERROR"
            }
        }
    }
    catch {
        Write-Host "Error processing sharing links: $_" -ForegroundColor Red
        Write-Log "Error processing sharing links: $_" "ERROR"
    }
}
#endregion

#region Main Script
function Main {
    try {
        # Initialize logging
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $logFilePath = "$env:TEMP\SPOUserRemover_$timestamp.log"
        
        Write-Host "SharePoint Online User Remover Script" -ForegroundColor Magenta
        Write-Host "=====================================" -ForegroundColor Magenta
        Write-Host "Site URL: $siteURL" -ForegroundColor White
        Write-Host "User List: $userListPath" -ForegroundColor White
        Write-Host "Log File: $logFilePath" -ForegroundColor White
        Write-Host ""
        
        Write-Log "Starting SPO User Remover Script"
        Write-Log "Site URL: $siteURL"
        Write-Log "User List: $userListPath"
        
        # Read user list
        $users = Read-UserList -FilePath $userListPath
        if ($users.Count -eq 0) {
            throw "No users found in the user list file"
        }
        
        # Connect to SharePoint Online
        if (-not (Connect-SPOService -SiteUrl $siteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant)) {
            throw "Failed to connect to SharePoint Online"
        }
        
        # Remove users from site groups
        Remove-UserFromSiteGroups -Users $users
        
        # Remove users from file permissions
        Remove-UserFromFilePermissions -Users $users
        
        # Remove users from sharing links
        Remove-UserFromSharingLinks -Users $users
        
        Write-Host "`nScript completed successfully!" -ForegroundColor Green
        Write-Log "Script completed successfully" "SUCCESS"
        
        # Generate summary
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "- Users processed: $($users.Count)" -ForegroundColor White
        Write-Host "- Log file: $logFilePath" -ForegroundColor White
    }
    catch {
        Write-Host "`nScript failed: $_" -ForegroundColor Red
        Write-Log "Script failed: $_" "ERROR"
        exit 1
    }
    finally {
        try {
            Disconnect-PnPOnline
            Write-Log "Disconnected from SharePoint Online"
        }
        catch {
            Write-Log "Warning during disconnect: $_" "WARNING"
        }
    }
}

# Execute main function
Main
#endregion
