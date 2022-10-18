find-module ImportExcel
install-module ImportExcel
Import-Module ImportExcel
Add-Type -Assembly  System.Windows.Forms

########################    GUI   #########################
$formObject = [System.Windows.Forms.Form]
$labelObject = [System.Windows.Forms.Label]
$buttonObject = [System.Windows.Forms.Button]
$menuBar = [System.Windows.Forms.MenuStrip]
$toolStrip = [System.Windows.Forms.ToolStrip]
$toolStripItem = [System.Windows.Forms.ToolStripMenuItem]
$toolStripButton = [System.Windows.Forms.ToolStripButton]
$listBox = [System.Windows.Forms.ListBox]
$comboBox = [System.Windows.Forms.ComboBox]
$listView = [System.Windows.Forms.ListView]

$baseForm = New-Object $formObject
$baseForm.height = 600
$baseForm.width = 1000
$baseForm.backColor = "white"
$baseForm.Text = 'Marine Corps Automated Distro Scrub'
$baseForm.AutoSize = $true

$aboutText = New-Object $labelObject
$aboutText.Text = "     Created and built by Sgt Schissler from 1st Intelligence Battalion. Written using powershell and .NET. 
`n      This GUI loads an excel sheet full of names, loads a text file with a list of Active Directory groups,
and iterates through first and last names within the excel roster to look for discrepancies within the groups. Used to solve the common time-wasting operation of 'distro scrubbing'.
`n      To run correctly, user will need to be using a powershell terminal with administrative permissions. Additionally, user must have permissions for the groups that he/she is trying to edit. 
`n      GUI was completely written by hand to give me practice with GUI programming in C#.
`n      Searching algorithm will be changed to 'binary search' in the future to reduce time complexity.
`n      Last updated on 2022-08-10"
$aboutText.Size = "400, 450"
$aboutText.Font = 'Arial,12,style=regular'
$aboutText.location = '10, 20'

$aboutForm = new-object $formObject
$aboutForm.height = 450
$aboutForm.width = 400
$aboutForm.backColor = "white"
$aboutForm.Text = 'Marine Corps Automated Distro Scrub'
$aboutForm.AutoSize = $true
$aboutForm.StartPosition = "CenterScreen"


$excelList = New-Object $listBox
$excelList.Location = "10, 100"
$excelList.Size = "250, 400"
$excelList.DisplayMember = 'FormattedName'
$excelList.Sorted = $true;

$excelListType = new-object $labelObject
$excelListType.Location = "100, 80"
$excelListType.Text = "(excel)"
$excelListType.Size = "150, 50"
$excelListType.Font = 'Arial,12,style=regular'

$excelListTitle = new-object $labelObject
$excelListTitle.Location = "50, 60"
$excelListTitle.Text = "S1 Master Roster"
$excelListTitle.Size = "200, 20"
$excelListTitle.Font = 'Arial,14,style=bold'

$ADMembersListTitle = new-object $labelObject
$ADMembersListTitle.Location = "550, 60"
$ADMembersListTitle.Text = "AD Group Members"
$ADMembersListTitle.Size = "200, 20"
$ADMembersListTitle.Font = 'Arial,14,style=bold'

$ADGroupDropdownTitle = new-object $labelObject
$ADGroupDropdownTitle.Location = "325, 60"
$ADGroupDropdownTitle.Text = "AD Group List"
$ADGroupDropdownTitle.Size = "200, 20"
$ADGroupDropdownTitle.Font = 'Arial,14,style=bold'

$lookupADTitle = new-object $labelObject
$lookupADTitle.Location = "270, 200"
$lookupADTitle.Text = "Load AD Group Selection"
$lookupADTitle.Size = "240, 50"
$lookupADTitle.Font = 'Arial,14,style=bold'

#region 
$ADInvalidGroupTitle = new-object $labelObject
$ADInvalidGroupTitle.Location = "825, 60"
$ADInvalidGroupTitle.Text = "Removable Members"
$ADInvalidGroupTitle.Size = "260, 50"
$ADInvalidGroupTitle.Font = 'Arial,14,style=bold'
#endregion

$ADIdentifierList = New-Object $comboBox
$ADIdentifierList.location = "280, 100"
$ADIdentifierList.Width = "225"

$ADGroupListView = New-Object $listView
$ADGroupListView.Location = "520, 100"
$ADGroupListView.Size = "250, 400"
$ADGroupListView.View = 'details'
$col1 = $ADGroupListView.Columns.Add("Display Name", -2)
$ADGroupListView.FullRowSelect = $true

$ADInvalidGroupListBox = New-Object $listBox
$ADInvalidGroupListBox.Location = "800, 100"
$ADInvalidGroupListBox.Size = "250, 400"
$ADInvalidGroupListBox.DisplayMember = 'DisplayName'
$ADInvalidGroupListBox.Sorted = $true;

$ADGroupListProcessButton = New-Object $buttonObject
$ADGroupListProcessButton.Text = 'Load Selection'
$ADGroupListProcessButton.Font = 'Arial,12,style=regular'
$ADGroupListProcessButton.Location = "300, 235"
$ADGroupListProcessButton.Size = "175, 30"

$searchGroupMembersTitle = new-object $labelObject
$searchGroupMembersTitle.Location = "270, 300"
$searchGroupMembersTitle.Text = "Search Group Members"
$searchGroupMembersTitle.Size = "240, 35"
$searchGroupMembersTitle.Font = 'Arial,14,style=bold'

$searchGroupMembersButton = New-Object $buttonObject
$searchGroupMembersButton.Text = 'Search'
$searchGroupMembersButton.Font = 'Arial,12,style=regular'
$searchGroupMembersButton.Location = "300, 335"
$searchGroupMembersButton.Size = "175, 30"

#toolbar
$mainMenu = New-Object $menuBar
$mainToolStrip = New-Object $toolStrip

$menuFile = New-Object $toolStripItem
$menuFile.Text = "File"

$menuOpen = New-Object $toolStripItem
$menuOpen.Text = "Open Master Roster (Excel)"

$menuAbout = New-Object $toolStripItem
$menuAbout.text = "About"

$menuOpenADIdentifiers = New-Object $toolStripItem
$menuOpenADIdentifiers.Text = "Open AD Group List (Text File)"

$toolStripOpen = New-Object $toolStripButton

$baseForm.MainMenuStrip = $mainMenu






#array to draw components
$baseForm.Controls.AddRange(@($mainMenu, $excelList, $excelListTitle, $excelListType, $ADIdentifierList, 
$ADGroupListProcessButton, $ADGroupDropdownTitle, $lookupADTitle, $ADMembersListTitle, $searchGroupMembersTitle, $searchGroupMembersButton,
$ADGroupListView, $ADInvalidGroupListBox, $ADInvalidGroupTitle))
[void]$baseForm.Controls.Add($mainToolStrip)
$mainToolStrip.Items.Add($menuFile)
$mainToolStrip.Items.Add($menuAbout)
[void]$menuFile.DropDownItems.Add($menuOpen)
[void]$menuFile.DropDownItems.Add($menuOpenADIdentifiers)


function open-File()
{
    $locateExcelSheet = New-Object System.Windows.Forms.OpenFileDialog 
    $locateExcelSheet.InitialDirectory = "C:\\"
    $locateExcelSheet.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $locateExcelSheet.RestoreDirectory = $true | Out-Null

    $response = $locateExcelSheet.ShowDialog() 

    if($response -eq 'Ok')  
    { 
        $excelList.Items.clear()
    }
    return $locateExcelSheet.FileName
}

function loadAbout()
{
    $aboutForm.Controls.Add($aboutText)
    $aboutForm.ShowDialog()

}

function loadExcelValues
{   
    param
    (
        [Parameter(ValueFromPipeline=$true)]
        $excelLocation
    )
    # Parameter help description
    
    if($excelLocation.length -gt 0)
    {
        $users = Import-Excel -Path $excelLocation -HeaderName 'Rank', 'Name' -StartRow 2
    }
    else
    {
        "User cancelled file opening" | write-host
        return
    }
    foreach($user in $users)
    {
        #creates new PSOBject per user
        $userLoader = New-Object -TypeName psobject
     
        #Begins process to split up fullnames from excel into pieces that can be searchable
        [string]$fullName  = $user.Name

        #check if user has a middle name. If not, perform abreviated operation
        if($fullName.Contains('.') -eq $false)
        {
            $nameSplitter = @($fullname.Split(','))
 
            Add-Member -InputObject $userLoader -MemberType NoteProperty -Name 'LName' -Value $nameSplitter[0]
            Add-Member -InputObject $userLoader -MemberType NoteProperty -Name 'FName' -Value $nameSplitter[1].Trim(' ')
            Add-Member -InputObject $userLoader -MemberType NoteProperty -Name 'Rank'  -Value $user.Rank

            $FormattedNameString = $userLoader.LName + " " + $userLoader.Rank + " " + $userLoader.FName
            Add-Member -InputObject $userLoader -MemberType NoteProperty -Name 'FormattedName' -Value $FormattedNameString

            [void]$excelList.Items.Add($userLoader)

        }

        #if user does has a middle name, perform normal operation
        else 
        {
            $nameSplitter = @($fullname.Split(','))
            [string]$middleInitial = $nameSplitter[1]
            if($middleInitial.length -gt 2)
            {
                $middleInitial = $middleInitial.Substring($middleInitial.Length - 2)
            }

            [string]$firstName = $nameSplitter[1]
            $firstName = $firstName.Substring(0, $firstName.Length - 3)
            $firstName = $firstName.Substring(1)

            #Add pieces of split array to temporary object $userLoader to act as searchable properties
            Add-Member -InputObject $userLoader -MemberType NoteProperty -Name 'LName' -Value $nameSplitter[0]
            Add-Member -InputObject $userLoader -MemberType NoteProperty -Name 'MInitial' -Value $middleInitial
            Add-Member -InputObject $userLoader -MemberType NoteProperty -Name 'FName' -Value $firstName
            Add-Member -InputObject $userLoader -MemberType NoteProperty -Name 'Rank'  -Value $user.Rank

            #Create a formatted String of the object's properties to look nice in the GUI
            $FormattedNameString = $userLoader.LName + " " + $userLoader.Rank + " " + $userLoader.FName + ' ' + $userLoader.MInitial
            Add-Member -InputObject $userLoader -MemberType NoteProperty -Name 'FormattedName' -Value $FormattedNameString

            [void]$excelList.Items.Add($userLoader)
            
        }
        
        
    }
}

function compareRosterToGroup()
{

    $indexesOfInvalidMembers = @()
    $invalidMembers = New-Object -TypeName "System.Collections.ArrayList"
    $groupMemberIndex = 0
    #iterate through each member in the group list
    if($ADgroupListView.Items.Count -eq 0)
    {
        return $null
    }
    foreach($groupMember in $ADGroupListView.Items)
    {
        #need to check to see if the account exists in the excel group
        [int]$iterator = 0 
        $foundMember = $false
        while($iterator -lt $excelList.Items.Count)
        {
            if(($groupMember.Tag.surname -eq $excelList.items[$iterator].LName) -and ($groupMember.Tag.givenName -eq $excelList.items[$iterator].FName)) 
            {
                $foundMember = $true
                break
                
            }
            $iterator++
        }

        if(!$foundMember)
        {
            $invalidMembers.add($groupMember.Tag)
            $indexesOfInvalidMembers += $groupMemberIndex
            #$ADGroupListBox.Items[$iterator].text.backColor = 'red'
            #paint the object on the table
        }

        $groupMemberIndex++
    }

    #color invalid user listview items with red background
    
    if($excelList.Items.Count -gt 0)
    {
        foreach($x in $indexesOfInvalidMembers)
        {
            $ADGroupListView.Items.Item($x).backColor = 'red'
        }
    }
    try
    {
        $ADInvalidGroupListBox.Items.AddRange($invalidMembers.displayname)
    }
    catch
    {
        Write-Host "`nNo invalid members found"
        
    }
    return ,$invalidMembers
}



function clearComboBoxList()
{
    $ADIdentifierList.Items.Clear()
}

function loadDropDown
{
    param
    (
        [Parameter(ValueFromPipeline=$true)]
        $textFileLocation
    )
    clearComboBoxList
    
    try
    {
        $textFileStreamReader = new-object System.IO.StreamReader($textFileLocation)
    }
    catch
    {
        "User cancelled file opening" | write-host
        return
    }
    while($null -ne ($getCurrentLine = $textFileStreamReader.ReadLine()))
    {
        $ADIdentifierList.Items.Add($getCurrentLine)
    }
    
    try
    {
        $ADIdentifierList.SelectedIndex = 0
    }
    catch
    {
        [System.Windows.Forms.MessageBox]::Show("`n`nPlease select a valid text file to load", "Invalid Users", "Ok")
    }
    
}

function readInADGroupIdentifiers()
{
    $locateTextFile = New-Object System.Windows.Forms.OpenFileDialog 
    $locateTextFile.InitialDirectory = "C:\\"
    $locateTextFile.Filter = "Group Text Files (*txt)|*.txt"
    $locateTextFile.RestoreDirectory = $true | Out-Null
    $locateTextFile.ShowDialog() | out-null
    try
    {
        $groupIdentifiers = $locateTextFile.FileName
    }
    catch {
        #do nothing, user hit 'cancel' 
    }
    return $groupIdentifiers
}

function clearBoxList()
{
    $ADGroupListBox.Items.Clear()
}

function getADGroupMembers()
{
    $ADGroupListView.Items.Clear()
    $currentItem = $ADIdentifierList.SelectedItem
    if($null -eq $currentItem)
    {
        [System.Windows.Forms.MessageBox]::Show("No Active Directory group selected", "Invalid Users", "Ok")
        return
    }

    try
    {
        $getADGroupMemberObject = Get-ADGroupMember -Identity $currentItem
    }
    catch
    {

        return
    }
    
    try 
    {
        for($iterate = 0; $iterate -lt $getADGroupMemberObject.length; $iterate++)
        {
            $getADGroupMemberObject[$iterate] = Get-ADUser -Identity $getADGroupMemberObject[$iterate].samAccountName -Properties displayname

        }
        foreach($guy in $getADGroupMemberObject)
        {
             $listViewAdd = New-Object System.Windows.Forms.ListViewItem -ArgumentList $guy.displayname
             $listViewAdd.tag = $guy
             $ADGroupListView.Items.add($listViewAdd)
        }
    }
    catch 
    {
        $getADGroupMemberObject = Get-ADUser -Identity $getADGroupMemberObject.samAccountName -Properties displayname

        $listViewAdd = New-Object System.Windows.Forms.ListViewItem -ArgumentList $getADGroupMemberObject.displayname
        $listViewAdd.tag = $getADGroupMemberObject
        $ADGroupListView.Items.add($listViewAdd)
    }
}

function removeInvalidMembers()
{
    param
    (
        [Parameter(ValueFromPipeline=$true)]
        [System.Object[]]$invalidMembers
    )

    $invalidMemberCount =  $invalidMembers.Count
    $ADSelectedGroup = $ADIdentifierList.SelectedItem

    if(($excelList.Items.Count -eq 0) -and ($ADGroupListView.Items.Count -eq 0))
    {
        [System.Windows.Forms.MessageBox]::Show("No Master Roster loaded, and no group members loaded", "Invalid Users", "Ok")
        return
    }

    if($excelList.Items.Count -ne 0 -and ($ADGroupListView.Items.Count -eq 0))
    {
        [System.Windows.Forms.MessageBox]::Show("Group-list lookup failed, or group has no members", "Invalid Users", "Ok")
        return
    }

    if($excelList.Items.Count -eq 0 -and ($ADGroupListView.Items.Count -ne 0))
    {
        [System.Windows.Forms.MessageBox]::Show("No Master Roster loaded", "Invalid Users", "Ok")
        return
    }

    if($invalidMembers.Count -gt 0)
    {
        $response = [System.Windows.Forms.MessageBox]::Show("$invalidMemberCount user/s within group not found on Master Roster. Would you like to remove them from $ADSelectedGroup ?", "Invalid Users", "YesNo")

        if($response -eq 'Yes')
        {
            foreach($member in $invalidMembers)
            {
                try
                {
                    Remove-ADGroupMember -Identity $ADIdentifierList.SelectedItem -Members $member 
                    $memberGroups = Get-ADPrincipalGroupMembership -Identity $member.samAccountName | where{$_.name -eq $ADIdentifierList.SelectedItem}
                    if($null -eq $memberGroups)
                    {
                        write-host("`n`n" + $member.displayname + " was removed from the " + $ADIdentifierList.SelectedItem + " group`n")
                    }
                    getADGroupMembers  
                }
                catch
                {
                   [System.Windows.Forms.MessageBox]::Show("Powershell not being run with administrator permissions, or user cancelled operation. Please close, run powershell as an administrator, and run the script again to perform scrubbing", "Invalid Users", "Ok")
                   getADGroupMembers
                }    
            }
        }
        else 
        {
            for($x = 0; $x -lt $ADGroupListView.Items.Count; $x++)
            {
                $ADGroupListView.Items.Item($x).backColor = 'white'
            }
        }
    }
    else
    {
        
        [System.Windows.Forms.MessageBox]::Show("No user discrepancies found between Master Roster and the Active Directory Group", "Invalid Users", "Ok")
        for($x = 0; $x -lt $ADGroupListView.Items.Count; $x++)
        {
            $ADGroupListView.Items.Item($x).backColor = 'white'
        }
    }
    $ADInvalidGroupListBox.Items.Clear()
}


##add in functions

[void]$menuAbout.Add_Click({loadAbout})
[void]$menuOpen.Add_Click({open-File | loadExcelValues | Out-Null})
$ADGroupListProcessButton.Add_Click({getADGroupMembers})
$menuOpenADIdentifiers.Add_Click({readInADGroupIdentifiers | loadDropDown})
$searchGroupMembersButton.Add_Click({compareRosterToGroup | removeInvalidMembers})


$baseForm.ShowDialog()



$baseForm.Dispose();
$excelList.Dispose();
$ADGroupListView.Dispose();
$ADIdentifierList.Dispose();











































