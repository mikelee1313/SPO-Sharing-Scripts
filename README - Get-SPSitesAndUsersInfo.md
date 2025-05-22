

**Get-SPSitesAndUsersInfo.ps1**

This script connects to a SharePoint tenant using PnP PowerShell with certificate-based authentication.  It retrieves a list of SharePoint sites either from a provided CSV file or directly from the tenant. 

The script collects and consolidates a variety of information into the output CSV file. Below is the list of columns that are included the output file:


**API Requirements**

- Microsoft Graph| Application |  Directory.Read.All 
- SharePoint |Application | Sites.FullControl.All

Example:

![image](https://github.com/user-attachments/assets/be97b24e-f6c6-470e-9b3e-05666159a4c0)


**Site Information**
- URL: The URL of the SharePoint site.
- Owner: The owner of the site.
- IB Mode: Information Barrier (IB) mode applied to the site.
- IB Segment: The Information Barrier segments associated with the site.
- Group ID: The ID of the associated Microsoft 365 group.
- RelatedGroupId: The ID of a related Microsoft 365 group, if any.
- IsHubSite: Indicates if the site is a hub site.
- Template: The site template used.
- Community Site: Indicates if the site is a community site (e.g., Yammer-linked).
- Custom Script Allowed: Indicates if custom scripts are allowed on the site.
- IsTeamsConnected: Indicates if the site is connected to Microsoft Teams.
- IsTeamsChannelConnected: Indicates if the site is connected to a Teams channel.
- TeamsChannelType: Specifies the type of Teams channel connected.
- StorageQuota (MB): The storage quota allocated to the site, in megabytes.
- StorageUsageCurrent (MB): The current storage usage of the site, in megabytes.
- LockState: The current lock state of the site.
- LastContentModifiedDate: The date when the site's content was last modified.
- ArchiveState: The archive state of the site.

**Site Level Sharing Indicators**
- AllowMembersEditMembership: Indicates if members can edit group membership.
- MembersCanShare: Indicates if members can share content.
- Has Sharing Links: Indicates if sharing links are being used.
- SiteDefinedSharingCapability: Sharing settings defined at the site level.
- SharingCapability: The overall sharing capability of the site.
- DisableCompanyWideSharingLinks: Whether company-wide sharing links are disabled at the Site level
- EEEU Present: A flag indicating if the "Everyone Except External Users" group is present.
  
 **Version Policy**
- DefaultTrimMode: Default trimming mode for versioning.
- DefaultExpireAfterDays: Number of days after which content expires by default.
- MajorVersionLimit: The maximum number of major versions retained.

**Microsoft 365 Group Details**
- Entra Group Alias: Alias of the associated Microsoft 365 (Entra) group.
- Entra Group AccessType: Access type of the Microsoft 365 group.
- Entra Group WhenCreated: Creation date of the Microsoft 365 group.

**Users and groups in the Site**
- SP Groups On Site: List of all SharePoint groups on the site.
- SP Groups Roles: Roles assigned to each SharePoint group.
- Site Collection Admins (Name <Email>): Site collection administrators with their names and emails.
- SP Users (Group: Name <Email>): SharePoint users with their groups, names, and emails.
- Site Level Users (Name <Email> [Roles]): Site-level users with their roles.
- Entra Group Owners (Name <Email>): Owners of the associated Microsoft 365 group.
- Entra Group Members (Name <Email>): Members of the associated Microsoft 365 group.


  
**Output Examples:**

![image](https://github.com/user-attachments/assets/de35fea2-496f-4831-bb1f-a626808e6269)

![image](https://github.com/user-attachments/assets/80fc90c2-dab6-4f39-8866-6377ff2894e4)

![image](https://github.com/user-attachments/assets/d643448d-8bbc-4ec5-85cb-de08301332e5)

![image](https://github.com/user-attachments/assets/9dccd5f1-1977-4e16-b1b4-e305153a9560)

![image](https://github.com/user-attachments/assets/1b04cdd8-f14b-4011-ad20-7c794a175412)


Future Function:
Sharing link / EEEU total counts per site collection
Support for nested Groups
Support For Subsets


