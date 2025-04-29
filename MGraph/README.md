# MGraph

This directory contains new- or rewritten scripts using the [Microsoft Graph](https://learn.microsoft.com/en-us/graph/overview) modules.  
Each script will check if the current environment's permissions and will request additional permissions if required before running.

For additional information on each script, they each have [comment based help](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-7.5).

## LicenseReport.ps1

Script meant for reporting license usage on your tenant.  
Using the `-SkipTrial` switch will skip the license named "Trial" or "Free".

## NameCleanup.ps1

Short script meant for cleaning up the Entra ID and standardizing the user's display name.

The configured changes are:

- Reverse `<LastName>.<FirstName>` to `<FirstName> <LastName>`
- Moving titles to the front if the DisplayName and removing them from the surname
- Moving text between brackets to the end of th DisplayName and removing it from the currently associated field ('surName' or 'givenName')

## TeamsStructure.ps1

Script generates a csv-file containing the following information:

- GroupId : Id of the group associated to the Teams
- Teams : Name of the group/Teams
- TeamType : Public of private
- ChannelId : Id of the channel
- Channel : Name of the channel
- ChannelType : Public or private
