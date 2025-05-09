

**Get-SPSitesAndUsersInfo.ps1**

This script connects to a SharePoint tenant using PnP PowerShell with certificate-based authentication.  It retrieves a list of SharePoint sites either from a provided CSV file or directly from the tenant. 

The script collects and consolidates a variety of information into the output CSV file. Below is the list of columns that are included the output file:


**API Requirements**

- Microsoft Graph| Application |  Directory.Read.All 
- SharePoint |Application | Sites.FullControl.All

Example:

![image](https://github.com/user-attachments/assets/be97b24e-f6c6-470e-9b3e-05666159a4c0)


**Site Information**
- URL - The URL of the SharePoint site.
- Owner - The owner of the site.
- IB Mode - Information Barrier (IB) mode for the site.
- IB Segment - Information Barrier segments for the site.
- Group ID - Microsoft 365 Group ID associated with the site (if applicable).
- RelatedGroupId - Related Group ID for the site (if applicable).
- IsHubSite - Indicates if the site is a Hub Site.
- Template - The template used for the site.
- SiteDefinedSharingCapability - Site-defined sharing capability.
- SharingCapability - Overall sharing capability for the site.
- DisableCompanyWideSharingLinks - Indicates if company-wide sharing links are disabled.
- Custom Script Allowed - Indicates if custom scripts are allowed.
- IsTeamsConnected - Indicates if the site is connected to Microsoft Teams.
- IsTeamsChannelConnected - Indicates if the site has a connected Teams channel.
- TeamsChannelType - Type of Teams channel connected to the site.
- StorageQuota (MB) - Total storage quota of the site in MB.
- StorageUsageCurrent (MB) - Current storage usage of the site in MB.
- LockState - Lock state of the site.
- LastContentModifiedDate - Last content modification date of the site.
- ArchiveState - Archive state of the site.

**Version Policy**
- DefaultTrimMode - Default trim mode for versioning.
- DefaultExpireAfterDays - The lifespan of items before they expire (if set).
- MajorVersionLimit - Limit on the number of major versions stored.

**Microsoft 365 Group Details**
- Entra Group Displayname - Display name of the associated Entra (Microsoft 365) Group.
- Entra Group Alias - Alias of the associated Microsoft 365 Group.
- Entra Group AccessType - Access type (e.g., public or private) of the Microsoft 365 Group.
- Entra Group WhenCreated - Creation date of the Microsoft 365 Group.

**Site Collection Administrators**
- Site Collection Admins (Name <Email>) - List of site collection administrators in the format "Name <Email>".

**Sharing Indicators**
- Has Sharing Links - Indicates if there are sharing links.
- Shared With Everyone - Indicates if the site is shared with "Everyone."

**SharePoint Groups**
- SP Groups On Site - List of SharePoint groups on the site.
- SP Groups Roles - Roles assigned to the SharePoint groups.

**Site Users**
- SP Users (Group: Name <Email>) - List of SharePoint users, grouped by associated SharePoint groups, in the format "Group:Name <Email>".
- Entra Group Owners (Name <Email>) - List of Microsoft 365 Group owners in the format "Name <Email>".
- Entra Group Members (Name <Email>) - List of Microsoft 365 Group members in the format "Name <Email>".

  
**Output Examples:**

![image](https://github.com/user-attachments/assets/de35fea2-496f-4831-bb1f-a626808e6269)

![image](https://github.com/user-attachments/assets/80fc90c2-dab6-4f39-8866-6377ff2894e4)

![image](https://github.com/user-attachments/assets/d643448d-8bbc-4ec5-85cb-de08301332e5)

![image](https://github.com/user-attachments/assets/9dccd5f1-1977-4e16-b1b4-e305153a9560)

![image](https://github.com/user-attachments/assets/1b04cdd8-f14b-4011-ad20-7c794a175412)


Future Function:
Collect Root users and Permissions


