# QuickScrub
PowerShell script to scrub through USMC Active Directory groups to remove members not currently on the unit's "Master Personnel Roster".

Requires local administrative priviledges to run the script correctly, as well as permissions on the current domain to access and delete members from AD-Groups.

First, load the provided excel document with the unit's master roster list, denoted by rank and then name. Then, load a list of AD Groups that you wish to parse through into the .txt file, each group presenting on a new-line.

Then, run the powershell script. Utilizing the "File" button in the top left corner, load both the Master Roster as well as the AD Group List into the program. Then, utilizing the "Load Selection" and "Search" buttons, the script will parse through the members in the groups to look for invalid users. If it finds any, the users will be added to a window on the far right side, and will become highlighted in red compared to the other groups members. A window will also appear and prompt the Administrator to confirm removal of these members. The Administrator will then have to open up the PowerShell command prompt again to further verify and confirm removal of group members from the groups in order to prevent mistakes. 
