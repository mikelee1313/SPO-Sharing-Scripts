# Get-SPOSiteSharingLinks

The main purpose of this script is to read the Sharing Links groups at the site collection level, to identify potential internal oversharing.

This script will use multiple site inputs, like "Get-SPOSite" filtering OneDrive or SPO Sites or a list of URLs from a site list.

**Important:** The main loop is using Get-SPOSiteGroup to loop through all SharePoint Site groups and users. However, you need to be a Site Collection Admin to read this data, so please be aware that the script will add your specified account as a Site Collection Admin to the site, get the Site Groups, the remove the account as Site Collection Admin.

The script will also collect several Site and Group properties that can used to catalog and understand groupings of sites. For example,  “non-group Connected Sites, Group Connected sites, Sites with Teams, Teams private channels and etc.

Default properties collected:

Site Props:
"URL"
"Owner" 
"IB Mode"
"IB Segment"
"Group ID" 
"RelatedGroupId"
"IsHubSite"
"Template"
"SiteDefinedSharingCapability"
"SharingCapability"
"IsTeamsConnected"
"IsTeamsChannelConnected"
"TeamsChannelType"
  

Group Props:
"SPGroup Title"
"SPGroup Roles"
"SPGroup Users"
"Entra Group Displayname"
"Entra Group Alias"
"Entra Group AccessType"
"Entra Group ManagedBy"
"Entra Group WhenCreated"
"Entra Group Owners"
"Entra Group Members"

If you have Information Barriers in your Teant, you will see the Information Barriers Segments and Information Barriers mode at the site collection levels.


![image](https://github.com/mikelee1313/Get-SPOSiteSharingLinks/assets/62190454/5fa98621-4594-4c7d-a39a-671ded1387af)


More information regarding Organization Wide Sharing Links

Here are 3 sharable link types:

•	People in [your organization]:  Gives anyone in your organization who has the link access to the file, whether they receive it directly from you or forwarded from someone else.
•	People you choose gives access only to the people you specify, although other people may already have access. If people forward the sharing invitation, only people who already have access to the item will be able to use the link.  
•	Anyone:  Gives access to anyone who receives this link, whether they receive it directly from you or forwarded from someone else. This may include people outside of your organization.

Example of sharing a file with everyone in the organization (Organization Links):

![image](https://github.com/mikelee1313/Get-SPOSiteSharingLinks/assets/62190454/a2de2d50-0c73-40b5-9089-282cf7d386d9)

When creating Organization Links at the file level, a Site level group is created with a name that is like “SharingLinks.296ee70c-2c8f-466e-bb5d-45f8eb805b1f.OrganizationEdit.96fe2ee1-486b-41fe-a467-e2aa4b102d11”

**Group Parts:**
SharingLinks.296ee70c-2c8f-466e-bb5d-45f8eb805b1f.OrganizationEdit.96fe2ee1-486b-41fe-a467-e2aa4b102d11
Link Name.DocumentID.LinkType.GroupID

Example:

![image](https://github.com/mikelee1313/Get-SPOSiteSharingLinks/assets/62190454/e2e05093-2a48-49a1-99d7-84b0205b1280)

**OrganizationEdit** = Company Wide links:
**AnonymousEdit** = Anyone Links
**Flexible**: People you choose

When using Org-wide sharing links,  you will see from above that the “SPGroup Users” only list a few users. These are the users that clicked the link.  Once a new user clicks the sharing link, they are dynamically added the “SPGroup Users” group. If a user is not listed this group, they will not automatically have access to the file and Copilot will not be able to discover this data during prompt response.
