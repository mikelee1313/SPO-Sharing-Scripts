# SPO SharingLinks Information

Lots of customers are in the process of getting ready for Copilot and identifying internal oversharing is a requirement for Copilot readiness and deployment. The main purpose of this script is to read the Sharing Links groups at the site collection level, to help customers identify potential internal oversharing in their tenant.


Here are 3 sharable link types:

•	**People in [your organization]**:  Gives anyone in your organization who has the link access to the file, whether they receive it directly from you or forwarded from someone else.

•	**People you choose**: gives access only to the people you specify, although other people may already have access. If people forward the sharing invitation, only people who already have access to the item will be able to use the link.  

•	**Anyone: ** Gives access to anyone who receives this link, whether they receive it directly from you or forwarded from someone else. This may include people outside of your organization.

Example of sharing a file with everyone in the organization (Organization Links):

![image](https://github.com/mikelee1313/Get-SPOSiteSharingLinks/assets/62190454/a2de2d50-0c73-40b5-9089-282cf7d386d9)

When creating Organization Links at the file level, a Site level group is created with a name that is like “SharingLinks.296ee70c-2c8f-466e-bb5d-45f8eb805b1f.OrganizationEdit.96fe2ee1-486b-41fe-a467-e2aa4b102d11”

**Group Parts:**

**Link Name.DocumentID.LinkType.GroupID**

SharingLinks.296ee70c-2c8f-466e-bb5d-45f8eb805b1f.OrganizationEdit.96fe2ee1-486b-41fe-a467-e2aa4b102d11


Example:

![image](https://github.com/mikelee1313/Get-SPOSiteSharingLinks/assets/62190454/e2e05093-2a48-49a1-99d7-84b0205b1280)

**OrganizationEdit** = Company Wide links:
**AnonymousEdit** = Anyone Links
**Flexible**: People you choose

Key Takeways:

•	When viewing Org-wide sharing links, you may see a subset of users in the “SPGroup Users” column. These are the users that clicked (redeemed) the link. 

•	Once a user clicks the sharing link (redeems the invitation), they are dynamically added the “SPGroup Users” group. 

•	Company wide link invitations can be redeemed by clicking the link, e-mailing the link, or dropping the link in a Teams chat.

•	If a user is not listed this group, they will not automatically have access to the file and Copilot will not be able to discover this data during prompt response.
