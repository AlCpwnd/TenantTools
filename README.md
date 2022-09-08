# TenantTools
Scripts meant to facilitate tenant administration.

## Teams Scripts
> **Disclaimer:**
> Neither of the following scripts are particularely efficient and runtime could be improved by using the Microsoft Graph module.
> I have yet to get around to implement it within the script. So depending on the amount of Teams your user has access to it might take > some time.

### Teams_AccessReport.ps1
Will return an array containing the current Teams and private channels the user has access to.

### Teams_AccessCopy.ps1
Will copy an existing user's access onto another existing user.

#### Parameters
- Template : Can accept a user UPN or the path towards a report generated with **Teams_AccessReport.ps1**.
- Identity : UPN of the user on which you want to apply the copied rights.
- IncludeRights : If enabled, the script will replicate the ownership of channels and Teams where applicable.
- Select : If enabled, you will be presented with a list of the template user's access prior to application. You can then select which access you want to replicate.

## Exchange Online Scripts

### Exch_MailForward.ps1
Allows a CSV with specific headers to be used to mass configure mail forwarding on mailboxes.

Csv headers:
- DisplayName: Strictly meant to for display purposes. This will be solely used in the `Write-Progress` part of the script.
- UserPrincipalName: UPN of the mailbox from which you want to forward the mails.
- Mail: Email address to which the mails will be forwarded.

## MSOL Scripts

### MSOL_ServiceCheck.ps1
Will go over all licensed users within the tenant and return any user that had services/features disabled from its assigned license.

#### Parameters
- Type: Will only accept the following 2 options
    - Standard: Will return a simple array of the users having the issue
    - Detailed: Will return a array containing the user, the affected license and the affected service(s) within said license