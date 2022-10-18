# QuickScrub
PowerShell script to scrub through USMC Active Directory groups to remove members not currently on the unit's "Master Personnel Roster".

Requires local administrative priviledges to run the script correctly, as well as permissions on the current domain to access and delete members from AD-Groups.

First, load the excel document with the unit's master roster list, denoted by rank and then name. Then, load a list of AD Groups that you wish to parse through into the 
