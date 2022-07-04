# TenantTools
Scripts meant to facilitate tenant administration.

## Teams Scripts
**Disclaimer:**
Neither of the following scrpipts are particularely efficient and runtime could be improved by using the Microsoft Graph module.
I have yet to get around to implement it within the script. So depending on the amount of Teams your user has access to it might take some time.

### Teams_AccessReport.ps1
Will return an array containing the current Teams and private channels the user has access to.

### Teams_AccessCopy.ps1
Will copy an existing user's access onto another existing user.

#### Parameters
- Template : Can accept a user UPN or the path towards a report generated with **Teams_AccessReport.ps1**.
- Identity : UPN of the user on which you want to apply the copied rights.
- IncludeRights : If enabled, the script will replicate the ownership of channels and Teams where applicable.
- Select : If enabled, you will be presented with a list of the template user's access prior to application. You can then select which access you want to replicate.