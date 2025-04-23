

**Get-SPSitesAndUsersInfo.ps1**

This script connects to a SharePoint tenant using PnP PowerShell with certificate-based authentication.  It retrieves a list of SharePoint sites either from a provided CSV file or directly from the tenant. 

For each site, it gathers comprehensive details including site properties, SharePoint groups and their roles, SharePoint users, Microsoft 365 group details (if applicable), group owners and members, and site collection administrators. 

The script consolidates this information into a structured format and exports it to a CSV file for reporting and auditing purposes.

Output Headers Collected:

URL	
Owner	
IB Mode	
IB Segment	
Group ID	
RelatedGroupId	
IsHubSite	
Template	
SiteDefinedSharingCapability	
SharingCapability	
DisableCompanyWideSharingLinks	
Custom Script Allowed	
IsTeamsConnected	
IsTeamsChannelConnected	
TeamsChannelType	
Entra Group Displayname	
Entra Group Alias	Entra Group 
AccessType	
Entra Group WhenCreated	
Site Collection Admins (Name <Email>)	
Has Sharing Links	
Shared With Everyone	
SP Groups On Site	
SP Groups Roles	
SP Users (Group: Name <Email>)	
Entra Group Owners (Name <Email>)	
Entra Group Members (Name <Email>)


  
Output Examples:

![image](https://github.com/user-attachments/assets/de35fea2-496f-4831-bb1f-a626808e6269)

![image](https://github.com/user-attachments/assets/80fc90c2-dab6-4f39-8866-6377ff2894e4)

![image](https://github.com/user-attachments/assets/d643448d-8bbc-4ec5-85cb-de08301332e5)

![image](https://github.com/user-attachments/assets/9dccd5f1-1977-4e16-b1b4-e305153a9560)

![image](https://github.com/user-attachments/assets/1b04cdd8-f14b-4011-ad20-7c794a175412)
