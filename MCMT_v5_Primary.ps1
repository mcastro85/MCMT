# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Initialize the domain name variable
$global:domainName = "Not Connected"

# Function to create the GUI
# Function to create the GUI
function New-GUI {
    # Create the main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "MCMT"
    $form.Size = New-Object System.Drawing.Size(1200, 800)
    $form.StartPosition = "CenterScreen"

    # Create the TabControl
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(1160, 650)
    $tabControl.Location = New-Object System.Drawing.Point(20, 100)


    # Create tabs
    $tabs = @("Computers", "Contacts", "Containers & OUs", "Groups", "Group Policy Objects", "Printers", "Users", "Search", "Processes/Services")

    foreach ($tabName in $tabs) {
        $tabPage = New-Object System.Windows.Forms.TabPage
        $tabPage.Text = $tabName
        $tabControl.TabPages.Add($tabPage)
    }

    $form.Controls.Add($tabControl)

    # Create buttons on the main GUI screen
    $newQueryButton = New-Object System.Windows.Forms.Button
    $newQueryButton.Text = "New Query"
    $newQueryButton.Size = New-Object System.Drawing.Size(150, 30)
    $newQueryButton.Location = New-Object System.Drawing.Point(20, 20)
    $newQueryButton.Add_Click({
        Show-NewQueryForm
    })
    $form.Controls.Add($newQueryButton)

    $networkButton = New-Object System.Windows.Forms.Button
    $networkButton.Text = "Network"
    $networkButton.Size = New-Object System.Drawing.Size(150, 30)
    $networkButton.Location = New-Object System.Drawing.Point(180, 20)
    $networkButton.Add_Click({
        Show-NetworkInfoForm
    })
    $form.Controls.Add($networkButton)

    $exportResultsButton = New-Object System.Windows.Forms.Button
    $exportResultsButton.Text = "Export Results"
    $exportResultsButton.Size = New-Object System.Drawing.Size(150, 30)
    $exportResultsButton.Location = New-Object System.Drawing.Point(340, 20)
    $form.Controls.Add($exportResultsButton)

    $generateCommandLineButton = New-Object System.Windows.Forms.Button
    $generateCommandLineButton.Text = "Generate Command Line"
    $generateCommandLineButton.Size = New-Object System.Drawing.Size(200, 30)
    $generateCommandLineButton.Location = New-Object System.Drawing.Point(500, 20)
    $form.Controls.Add($generateCommandLineButton)

    # Domain Settings Button
    $domainSettingsButton = New-Object System.Windows.Forms.Button
    $domainSettingsButton.Text = "Domain Settings"
    $domainSettingsButton.Size = New-Object System.Drawing.Size(150, 30)
    $domainSettingsButton.Location = New-Object System.Drawing.Point(710, 20)
    $domainSettingsButton.Add_Click({
        Show-DomainSettingsForm
    })
    $form.Controls.Add($domainSettingsButton)

    # Domain Name Label
    $domainNameLabel = New-Object System.Windows.Forms.Label
    $domainNameLabel.Text = "Domain: $global:domainName"
    $domainNameLabel.AutoSize = $true
    $domainNameLabel.Location = New-Object System.Drawing.Point(870, 25)
    $form.Controls.Add($domainNameLabel)

    # Azure Storage Button
    $azureStorageButton = New-Object System.Windows.Forms.Button
    $azureStorageButton.Text = "Azure Storage"
    $azureStorageButton.Size = New-Object System.Drawing.Size(150, 30)
    $azureStorageButton.Location = New-Object System.Drawing.Point(20, 60)
    $azureStorageButton.Add_Click({
        Show-AzureStorageForm
    })
    $form.Controls.Add($azureStorageButton)

    # CrowdStrike Button
    $crowdStrikeButton = New-Object System.Windows.Forms.Button
    $crowdStrikeButton.Text = "CrowdStrike"
    $crowdStrikeButton.Size = New-Object System.Drawing.Size(150, 30)
    $crowdStrikeButton.Location = New-Object System.Drawing.Point(180, 60)
    $crowdStrikeButton.Add_Click({
        Show-CrowdStrikeForm
    })
    $form.Controls.Add($crowdStrikeButton)

    # Computers Tab
    $computersTab = $tabControl.TabPages[0]
    Add-ADControls -parent $computersTab -type "Computer"

    # Contacts Tab
    $contactsTab = $tabControl.TabPages[1]
    Add-ContactsControls -parent $contactsTab

    # Containers & OUs Tab
    $containersOUsTab = $tabControl.TabPages[2]
    Add-ContainersControls -parent $containersOUsTab

    # Groups Tab
    $groupsTab = $tabControl.TabPages[3]
    Add-GroupsControls -parent $groupsTab

    # Group Policy Objects Tab
    $gpoTab = $tabControl.TabPages[4]
    Add-GPOControls -parent $gpoTab

    # Printers Tab
    $printersTab = $tabControl.TabPages[5]
    Add-PrintersControls -parent $printersTab

    # Users Tab
    $usersTab = $tabControl.TabPages[6]
    Add-UsersControls -parent $usersTab

    # Search Tab
    $searchTab = $tabControl.TabPages[7]
    Add-SearchControls -parent $searchTab

    # Processes/Services Tab
    $processesServicesTab = $tabControl.TabPages[8]
    Add-ProcessesServicesControls -parent $processesServicesTab

    # Show the form
    $form.Add_Shown({$form.Activate()})
    [void] $form.ShowDialog()
}

# Function to add AD controls to a tab
function Add-ADControls {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$parent,
        [Parameter(Mandatory=$true)][string]$type
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter $type Name:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $parent.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size(200, 20)
    $textBox.Location = New-Object System.Drawing.Point(150, 20)
    $parent.Controls.Add($textBox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Lookup"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(150, 50)
    $parent.Controls.Add($button)

    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"
    $resultBox.Size = New-Object System.Drawing.Size(1100, 400)
    $resultBox.Location = New-Object System.Drawing.Point(20, 90)
    $parent.Controls.Add($resultBox)

    $button.Add_Click({
        $query = switch ($type) {
            "Computer" {
                Get-ADComputer -Identity $textBox.Text -Properties * | Select-Object @{
                    Name = "Creation Date"
                    Expression = { $_.whenCreated }
                }, @{
                    Name = "Critical System Object"
                    Expression = { $_.isCriticalSystemObject }
                }, @{
                    Name = "Deleted"
                    Expression = { $_.isDeleted }
                }, @{
                    Name = "Description"
                    Expression = { $_.Description }
                }, @{
                    Name = "Disabled"
                    Expression = { $_.Enabled -eq $false }
                }, @{
                    Name = "Distinguished Name"
                    Expression = { $_.DistinguishedName }
                }, @{
                    Name = "Group Membership (All)"
                    Expression = { $_.MemberOf -join ", " }
                }, @{
                    Name = "Group Membership (Direct)"
                    Expression = { $_.MemberOf -join ", " }
                }, @{
                    Name = "Group Membership (Indirect/Nested)"
                    Expression = { $_.MemberOf -join ", " }
                }, @{
                    Name = "GUID"
                    Expression = { $_.ObjectGUID }
                }, @{
                    Name = "Last Known Location"
                    Expression = { $_.LastKnownParent }
                }, @{
                    Name = "Last Logon Date"
                    Expression = { $_.LastLogonDate }
                }, @{
                    Name = "Last Logon DC"
                    Expression = { $_.LastLogonDC }
                }, @{
                    Name = "Modification Date"
                    Expression = { $_.whenChanged }
                }, @{
                    Name = "Name"
                    Expression = { $_.Name }
                }, @{
                    Name = "Operating System"
                    Expression = { $_.OperatingSystem }
                }, @{
                    Name = "Parent Container"
                    Expression = { $_.CanonicalName }
                }, @{
                    Name = "Password Last Changed"
                    Expression = { $_.PasswordLastSet }
                }, @{
                    Name = "Primary Group"
                    Expression = { $_.PrimaryGroupID }
                }, @{
                    Name = "Service Pack"
                    Expression = { $_.OperatingSystemServicePack }
                }, @{
                    Name = "Show In Advanced View Only"
                    Expression = { $_.ShowInAdvancedViewOnly }
                }, @{
                    Name = "SID"
                    Expression = { $_.SID }
                } | Format-List | Out-String
            }
            "Contact" {
                Get-ADObject -LDAPFilter "(objectClass=contact)" -Filter { Name -eq $textBox.Text } -Properties * | Format-List | Out-String
            }
            "Container" {
                Get-ADOrganizationalUnit -Filter { Name -eq $textBox.Text } -Properties * | Format-List | Out-String
            }
            "Group" {
                Get-ADGroup -Identity $textBox.Text -Properties * | Format-List | Out-String
            }
            "GPO" {
                Get-GPO -Name $textBox.Text | Format-List | Out-String
            }
            "Printer" {
                Get-ADObject -LDAPFilter "(objectClass=printQueue)" -Filter { Name -eq $textBox.Text } -Properties * | Format-List | Out-String
            }
            "User" {
                Get-ADUser -Identity $textBox.Text -Properties * | Format-List | Out-String
            }
        }
        $resultBox.Text = $query
    })
}

# Function to add contacts controls to the Contacts tab
function Add-ContactsControls {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$parent
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select a Query:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $parent.Controls.Add($label)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Size = New-Object System.Drawing.Size(1100, 400)
    $listView.Location = New-Object System.Drawing.Point(20, 50)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.Columns.Add("Name", 550)
    $listView.Columns.Add("Type", 550)

    # Add items to the list view
    $queries = @(
        @{Name = "All contacts"; Type = "Contact"},
        @{Name = "Contact with specified Exchange alias"; Type = "Contact"},
        @{Name = "Contact with specified GUID"; Type = "Contact"},
        @{Name = "Contact with specified name"; Type = "Contact"},
        @{Name = "Contact with specified primary SMTP address"; Type = "Contact"},
        @{Name = "Contact with specified SMTP address"; Type = "Contact"},
        @{Name = "Contacts created in the last 30 days"; Type = "Contact"},
        @{Name = "Contacts created in the last X days"; Type = "Contact"},
        @{Name = "Contacts deleted in the last 30 days"; Type = "Contact"},
        @{Name = "Contacts in specified company"; Type = "Contact"},
        @{Name = "Contacts in specified department"; Type = "Contact"},
        @{Name = "Contacts modified in the last 30 days"; Type = "Contact"},
        @{Name = "Contacts modified in the last X days"; Type = "Contact"},
        @{Name = "Contacts that are direct members of specified group"; Type = "Contact"},
        @{Name = "Contacts that are directly or indirectly members of specified group"; Type = "Contact"},
        @{Name = "Contacts that are hidden from Exchange address lists"; Type = "Contact"},
        @{Name = "Contacts that are managers"; Type = "Contact"},
        @{Name = "Contacts that are not direct members of specified group"; Type = "Contact"},
        @{Name = "Contacts that are not directly or indirectly members of specified group"; Type = "Contact"},
        @{Name = "Contacts that are not protected from deletion"; Type = "Contact"},
        @{Name = "Contacts that are protected from deletion"; Type = "Contact"},
        @{Name = "Contacts with a manager"; Type = "Contact"},
        @{Name = "Contacts with no manager"; Type = "Contact"},
        @{Name = "Contacts with specified last name"; Type = "Contact"},
        @{Name = "Contacts with specified manager"; Type = "Contact"},
        @{Name = "Deleted contacts"; Type = "Contact"}
    )

    foreach ($query in $queries) {
        $item = New-Object System.Windows.Forms.ListViewItem($query.Name)
        $item.SubItems.Add($query.Type)
        $listView.Items.Add($item)
    }

    $parent.Controls.Add($listView)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Run Query"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(20, 470)
    $parent.Controls.Add($button)

    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"
    $resultBox.Size = New-Object System.Drawing.Size(1100, 150)
    $resultBox.Location = New-Object System.Drawing.Point(20, 510)
    $parent.Controls.Add($resultBox)

    $button.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedItem = $listView.SelectedItems[0]
            $resultBox.Text = switch ($selectedItem.Text) {
                "All contacts" { Get-ADObject -Filter { objectClass -eq "contact" } -Properties * | Format-List | Out-String }
                "Contact with specified Exchange alias" { "Query for Contact with specified Exchange alias" }
                "Contact with specified GUID" { "Query for Contact with specified GUID" }
                "Contact with specified name" { "Query for Contact with specified name" }
                "Contact with specified primary SMTP address" { "Query for Contact with specified primary SMTP address" }
                "Contact with specified SMTP address" { "Query for Contact with specified SMTP address" }
                "Contacts created in the last 30 days" { "Query for Contacts created in the last 30 days" }
                "Contacts created in the last X days" { "Query for Contacts created in the last X days" }
                "Contacts deleted in the last 30 days" { "Query for Contacts deleted in the last 30 days" }
                "Contacts in specified company" { "Query for Contacts in specified company" }
                "Contacts in specified department" { "Query for Contacts in specified department" }
                "Contacts modified in the last 30 days" { "Query for Contacts modified in the last 30 days" }
                "Contacts modified in the last X days" { "Query for Contacts modified in the last X days" }
                "Contacts that are direct members of specified group" { "Query for Contacts that are direct members of specified group" }
                "Contacts that are directly or indirectly members of specified group" { "Query for Contacts that are directly or indirectly members of specified group" }
                "Contacts that are hidden from Exchange address lists" { "Query for Contacts that are hidden from Exchange address lists" }
                "Contacts that are managers" { "Query for Contacts that are managers" }
                "Contacts that are not direct members of specified group" { "Query for Contacts that are not direct members of specified group" }
                "Contacts that are not directly or indirectly members of specified group" { "Query for Contacts that are not directly or indirectly members of specified group" }
                "Contacts that are not protected from deletion" { "Query for Contacts that are not protected from deletion" }
                "Contacts that are protected from deletion" { "Query for Contacts that are protected from deletion" }
                "Contacts with a manager" { "Query for Contacts with a manager" }
                "Contacts with no manager" { "Query for Contacts with no manager" }
                "Contacts with specified last name" { "Query for Contacts with specified last name" }
                "Contacts with specified manager" { "Query for Contacts with specified manager" }
                "Deleted contacts" { "Query for Deleted contacts" }
                default { "No query selected." }
            }
        }
    })
}

# Function to add controls to the Containers & OUs tab
function Add-ContainersControls {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$parent
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select a Query:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $parent.Controls.Add($label)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Size = New-Object System.Drawing.Size(1100, 400)
    $listView.Location = New-Object System.Drawing.Point(20, 50)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.Columns.Add("Name", 550)
    $listView.Columns.Add("Type", 550)

    # Add items to the list view
    $queries = @(
        @{Name = "All containers"; Type = "Container"},
        @{Name = "Container with specified GUID"; Type = "Container"},
        @{Name = "Containers created in the last 30 days"; Type = "Container"},
        @{Name = "Containers created in the last X days"; Type = "Container"},
        @{Name = "Containers deleted in the last 30 days"; Type = "Container"},
        @{Name = "Containers modified in the last 30 days"; Type = "Container"},
        @{Name = "Containers modified in the last X days"; Type = "Container"},
        @{Name = "Containers that are not protected from deletion"; Type = "Container"},
        @{Name = "Containers that are protected from deletion"; Type = "Container"},
        @{Name = "Containers with GPOs linked"; Type = "Container"},
        @{Name = "Containers with less than X child objects"; Type = "Container"},
        @{Name = "Containers with more than X child objects"; Type = "Container"},
        @{Name = "Containers with specified name"; Type = "Container"},
        @{Name = "Deleted containers"; Type = "Container"},
        @{Name = "Empty containers"; Type = "Container"}
    )

    foreach ($query in $queries) {
        $item = New-Object System.Windows.Forms.ListViewItem($query.Name)
        $item.SubItems.Add($query.Type)
        $listView.Items.Add($item)
    }

    $parent.Controls.Add($listView)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Run Query"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(20, 470)
    $parent.Controls.Add($button)

    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"
    $resultBox.Size = New-Object System.Drawing.Size(1100, 150)
    $resultBox.Location = New-Object System.Drawing.Point(20, 510)
    $parent.Controls.Add($resultBox)

    $button.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedItem = $listView.SelectedItems[0]
            $resultBox.Text = switch ($selectedItem.Text) {
                "All containers" { Get-ADOrganizationalUnit -Filter * -Properties * | Format-List | Out-String }
                "Container with specified GUID" { "Query for Container with specified GUID" }
                "Containers created in the last 30 days" { "Query for Containers created in the last 30 days" }
                "Containers created in the last X days" { "Query for Containers created in the last X days" }
                "Containers deleted in the last 30 days" { "Query for Containers deleted in the last 30 days" }
                "Containers modified in the last 30 days" { "Query for Containers modified in the last 30 days" }
                "Containers modified in the last X days" { "Query for Containers modified in the last X days" }
                "Containers that are not protected from deletion" { "Query for Containers that are not protected from deletion" }
                "Containers that are protected from deletion" { "Query for Containers that are protected from deletion" }
                "Containers with GPOs linked" { "Query for Containers with GPOs linked" }
                "Containers with less than X child objects" { "Query for Containers with less than X child objects" }
                "Containers with more than X child objects" { "Query for Containers with more than X child objects" }
                "Containers with specified name" { "Query for Containers with specified name" }
                "Deleted containers" { "Query for Deleted containers" }
                "Empty containers" { "Query for Empty containers" }
                default { "No query selected." }
            }
        }
    })
}

# Function to add controls to the Groups tab
function Add-GroupsControls {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$parent
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select a Query:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $parent.Controls.Add($label)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Size = New-Object System.Drawing.Size(1100, 400)
    $listView.Location = New-Object System.Drawing.Point(20, 50)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.Columns.Add("Name", 550)
    $listView.Columns.Add("Type", 550)

    # Add items to the list view
    $queries = @(
        @{Name = "All groups"; Type = "Group"},
        @{Name = "Deleted groups"; Type = "Group"},
        @{Name = "Distribution groups"; Type = "Group"},
        @{Name = "Domain Local groups"; Type = "Group"},
        @{Name = "Global groups"; Type = "Group"},
        @{Name = "Group with specified Exchange alias"; Type = "Group"},
        @{Name = "Group with specified GUID"; Type = "Group"},
        @{Name = "Group with specified name"; Type = "Group"},
        @{Name = "Group with specified primary SMTP address"; Type = "Group"},
        @{Name = "Group with specified SID"; Type = "Group"},
        @{Name = "Group with specified SMTP address"; Type = "Group"},
        @{Name = "Groups created in the last 30 days"; Type = "Group"},
        @{Name = "Groups created in the last X days"; Type = "Group"},
        @{Name = "Groups deleted in the last 30 days"; Type = "Group"},
        @{Name = "Groups modified in the last 30 days"; Type = "Group"},
        @{Name = "Groups modified in the last X days"; Type = "Group"},
        @{Name = "Groups that are direct members of specified group"; Type = "Group"},
        @{Name = "Groups that are directly or indirectly members of specified group"; Type = "Group"},
        @{Name = "Groups that are hidden from Exchange address lists"; Type = "Group"},
        @{Name = "Groups that are managers"; Type = "Group"},
        @{Name = "Groups that do not contain specified member"; Type = "Group"},
        @{Name = "Groups used as primary groups"; Type = "Group"},
        @{Name = "Groups with a manager"; Type = "Group"},
        @{Name = "Groups with at least 1 member"; Type = "Group"},
        @{Name = "Groups with less than X members"; Type = "Group"},
        @{Name = "Groups with more than X members"; Type = "Group"},
        @{Name = "Groups with no manager"; Type = "Group"},
        @{Name = "Groups with no members"; Type = "Group"},
        @{Name = "Groups with specified manager"; Type = "Group"},
        @{Name = "Groups with specified member"; Type = "Group"},
        @{Name = "Groups with specified SID in SID History"; Type = "Group"},
        @{Name = "Security groups"; Type = "Group"},
        @{Name = "Universal groups"; Type = "Group"}
    )

    foreach ($query in $queries) {
        $item = New-Object System.Windows.Forms.ListViewItem($query.Name)
        $item.SubItems.Add($query.Type)
        $listView.Items.Add($item)
    }

    $parent.Controls.Add($listView)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Run Query"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(20, 470)
    $parent.Controls.Add($button)

    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"
    $resultBox.Size = New-Object System.Drawing.Size(1100, 150)
    $resultBox.Location = New-Object System.Drawing.Point(20, 510)
    $parent.Controls.Add($resultBox)

    $button.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedItem = $listView.SelectedItems[0]
            $resultBox.Text = switch ($selectedItem.Text) {
                "All groups" { Get-ADGroup -Filter * -Properties * | Format-List | Out-String }
                "Deleted groups" { "Query for Deleted groups" }
                "Distribution groups" { "Query for Distribution groups" }
                "Domain Local groups" { "Query for Domain Local groups" }
                "Global groups" { "Query for Global groups" }
                "Group with specified Exchange alias" { "Query for Group with specified Exchange alias" }
                "Group with specified GUID" { "Query for Group with specified GUID" }
                "Group with specified name" { "Query for Group with specified name" }
                "Group with specified primary SMTP address" { "Query for Group with specified primary SMTP address" }
                "Group with specified SID" { "Query for Group with specified SID" }
                "Group with specified SMTP address" { "Query for Group with specified SMTP address" }
                "Groups created in the last 30 days" { "Query for Groups created in the last 30 days" }
                "Groups created in the last X days" { "Query for Groups created in the last X days" }
                "Groups deleted in the last 30 days" { "Query for Groups deleted in the last 30 days" }
                "Groups modified in the last 30 days" { "Query for Groups modified in the last 30 days" }
                "Groups modified in the last X days" { "Query for Groups modified in the last X days" }
                "Groups that are direct members of specified group" { "Query for Groups that are direct members of specified group" }
                "Groups that are directly or indirectly members of specified group" { "Query for Groups that are directly or indirectly members of specified group" }
                "Groups that are hidden from Exchange address lists" { "Query for Groups that are hidden from Exchange address lists" }
                "Groups that are managers" { "Query for Groups that are managers" }
                "Groups that do not contain specified member" { "Query for Groups that do not contain specified member" }
                "Groups used as primary groups" { "Query for Groups used as primary groups" }
                "Groups with a manager" { "Query for Groups with a manager" }
                "Groups with at least 1 member" { "Query for Groups with at least 1 member" }
                "Groups with less than X members" { "Query for Groups with less than X members" }
                "Groups with more than X members" { "Query for Groups with more than X members" }
                "Groups with no manager" { "Query for Groups with no manager" }
                "Groups with no members" { "Query for Groups with no members" }
                "Groups with specified manager" { "Query for Groups with specified manager" }
                "Groups with specified member" { "Query for Groups with specified member" }
                "Groups with specified SID in SID History" { "Query for Groups with specified SID in SID History" }
                "Security groups" { "Query for Security groups" }
                "Universal groups" { "Query for Universal groups" }
                default { "No query selected." }
            }
        }
    })
}

# Function to add controls to the Group Policy Objects tab
function Add-GPOControls {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$parent
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select a Query:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $parent.Controls.Add($label)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Size = New-Object System.Drawing.Size(1100, 400)
    $listView.Location = New-Object System.Drawing.Point(20, 50)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.Columns.Add("Name", 550)
    $listView.Columns.Add("Type", 550)

    # Add items to the list view
    $queries = @(
        @{Name = "All GPOs"; Type = "GPO"},
        @{Name = "Deleted GPOs"; Type = "GPO"},
        @{Name = "GPO with specified unique ID"; Type = "GPO"},
        @{Name = "GPOs created in the last 30 days"; Type = "GPO"},
        @{Name = "GPOs created in the last X days"; Type = "GPO"},
        @{Name = "GPOs deleted in the last 30 days"; Type = "GPO"},
        @{Name = "GPOs modified in the last 30 days"; Type = "GPO"},
        @{Name = "GPOs modified in the last X days"; Type = "GPO"},
        @{Name = "GPOs with all settings disabled"; Type = "GPO"},
        @{Name = "GPOs with all settings enabled"; Type = "GPO"},
        @{Name = "GPOs with computer settings disabled"; Type = "GPO"},
        @{Name = "GPOs with user settings disabled"; Type = "GPO"}
    )

    foreach ($query in $queries) {
        $item = New-Object System.Windows.Forms.ListViewItem($query.Name)
        $item.SubItems.Add($query.Type)
        $listView.Items.Add($item)
    }

    $parent.Controls.Add($listView)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Run Query"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(20, 470)
    $parent.Controls.Add($button)

    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"
    $resultBox.Size = New-Object System.Drawing.Size(1100, 150)
    $resultBox.Location = New-Object System.Drawing.Point(20, 510)
    $parent.Controls.Add($resultBox)

    $button.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedItem = $listView.SelectedItems[0]
            $resultBox.Text = switch ($selectedItem.Text) {
                "All GPOs" { Get-GPO -All | Format-List | Out-String }
                "Deleted GPOs" { "Query for Deleted GPOs" }
                "GPO with specified unique ID" { "Query for GPO with specified unique ID" }
                "GPOs created in the last 30 days" { "Query for GPOs created in the last 30 days" }
                "GPOs created in the last X days" { "Query for GPOs created in the last X days" }
                "GPOs deleted in the last 30 days" { "Query for GPOs deleted in the last 30 days" }
                "GPOs modified in the last 30 days" { "Query for GPOs modified in the last 30 days" }
                "GPOs modified in the last X days" { "Query for GPOs modified in the last X days" }
                "GPOs with all settings disabled" { "Query for GPOs with all settings disabled" }
                "GPOs with all settings enabled" { "Query for GPOs with all settings enabled" }
                "GPOs with computer settings disabled" { "Query for GPOs with computer settings disabled" }
                "GPOs with user settings disabled" { "Query for GPOs with user settings disabled" }
                default { "No query selected." }
            }
        }
    })
}

# Function to add controls to the Printers tab
function Add-PrintersControls {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$parent
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select a Query:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $parent.Controls.Add($label)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Size = New-Object System.Drawing.Size(1100, 400)
    $listView.Location = New-Object System.Drawing.Point(20, 50)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.Columns.Add("Name", 550)
    $listView.Columns.Add("Type", 550)

    # Add items to the list view
    $queries = @(
        @{Name = "All printers"; Type = "Printer"},
        @{Name = "Deleted printers"; Type = "Printer"},
        @{Name = "Printer with specified comment"; Type = "Printer"},
        @{Name = "Printer with specified driver name"; Type = "Printer"},
        @{Name = "Printer with specified GUID"; Type = "Printer"},
        @{Name = "Printer with specified local name"; Type = "Printer"},
        @{Name = "Printer with specified location"; Type = "Printer"},
        @{Name = "Printer with specified port name"; Type = "Printer"},
        @{Name = "Printer with specified share name"; Type = "Printer"},
        @{Name = "Printers created in the last 30 days"; Type = "Printer"},
        @{Name = "Printers created in the last X days"; Type = "Printer"},
        @{Name = "Printers deleted in the last 30 days"; Type = "Printer"},
        @{Name = "Printers modified in the last 30 days"; Type = "Printer"},
        @{Name = "Printers modified in the last X days"; Type = "Printer"},
        @{Name = "Printers shared from specified server"; Type = "Printer"}
    )

    foreach ($query in $queries) {
        $item = New-Object System.Windows.Forms.ListViewItem($query.Name)
        $item.SubItems.Add($query.Type)
        $listView.Items.Add($item)
    }

    $parent.Controls.Add($listView)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Run Query"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(20, 470)
    $parent.Controls.Add($button)

    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"
    $resultBox.Size = New-Object System.Drawing.Size(1100, 150)
    $resultBox.Location = New-Object System.Drawing.Point(20, 510)
    $parent.Controls.Add($resultBox)

    $button.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedItem = $listView.SelectedItems[0]
            $resultBox.Text = switch ($selectedItem.Text) {
                "All printers" { Get-ADObject -Filter { objectClass -eq "printQueue" } -Properties * | Format-List | Out-String }
                "Deleted printers" { "Query for Deleted printers" }
                "Printer with specified comment" { "Query for Printer with specified comment" }
                "Printer with specified driver name" { "Query for Printer with specified driver name" }
                "Printer with specified GUID" { "Query for Printer with specified GUID" }
                "Printer with specified local name" { "Query for Printer with specified local name" }
                "Printer with specified location" { "Query for Printer with specified location" }
                "Printer with specified port name" { "Query for Printer with specified port name" }
                "Printer with specified share name" { "Query for Printer with specified share name" }
                "Printers created in the last 30 days" { "Query for Printers created in the last 30 days" }
                "Printers created in the last X days" { "Query for Printers created in the last X days" }
                "Printers deleted in the last 30 days" { "Query for Printers deleted in the last 30 days" }
                "Printers modified in the last 30 days" { "Query for Printers modified in the last 30 days" }
                "Printers modified in the last X days" { "Query for Printers modified in the last X days" }
                "Printers shared from specified server" { "Query for Printers shared from specified server" }
                default { "No query selected." }
            }
        }
    })
}

# Function to add controls to the Users tab
function Add-UsersControls {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$parent
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select a Query:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $parent.Controls.Add($label)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Size = New-Object System.Drawing.Size(1100, 400)
    $listView.Location = New-Object System.Drawing.Point(20, 50)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.Columns.Add("Name", 550)
    $listView.Columns.Add("Type", 550)

    # Add items to the list view
    $queries = @(
        @{Name = "All users"; Type = "User"},
        @{Name = "Deleted users"; Type = "User"},
        @{Name = "Disabled users"; Type = "User"},
        @{Name = "Enabled users"; Type = "User"},
        @{Name = "User with specified Employee ID"; Type = "User"},
        @{Name = "User with specified Exchange alias"; Type = "User"},
        @{Name = "User with specified GUID"; Type = "User"},
        @{Name = "User with specified name"; Type = "User"},
        @{Name = "User with specified primary SMTP address"; Type = "User"},
        @{Name = "User with specified SID"; Type = "User"},
        @{Name = "User with specified SID in SID History"; Type = "User"},
        @{Name = "User with specified SMTP address"; Type = "User"},
        @{Name = "Users created in the last 30 days"; Type = "User"},
        @{Name = "Users created in the last X days"; Type = "User"},
        @{Name = "Users deleted in the last 30 days"; Type = "User"},
        @{Name = "Users in specified company"; Type = "User"},
        @{Name = "Users in specified department"; Type = "User"},
        @{Name = "Users modified in the last 30 days"; Type = "User"},
        @{Name = "Users modified in the last X days"; Type = "User"},
        @{Name = "Users not using mailbox size limit defaults"; Type = "User"},
        @{Name = "Users not using specified group as primary group"; Type = "User"},
        @{Name = "Users that are allowed to dial in"; Type = "User"},
        @{Name = "Users that are direct members of specified group"; Type = "User"},
        @{Name = "Users that are directly or indirectly members of specified group"; Type = "User"},
        @{Name = "Users that are not allowed to change password"; Type = "User"},
        @{Name = "Users that are not direct members of specified group"; Type = "User"},
        @{Name = "Users that are not directly or indirectly members of specified group"; Type = "User"},
        @{Name = "Users that are not protected from deletion"; Type = "User"},
        @{Name = "Users that are protected from deletion"; Type = "User"},
        @{Name = "Users that do not have an Outlook 2010 / Lync photo"; Type = "User"},
        @{Name = "Users that do not require a password"; Type = "User"},
        @{Name = "Users that last logged on by specified DC"; Type = "User"},
        @{Name = "Users that will expire in the next 30 days"; Type = "User"},
        @{Name = "Users where dial in permission is controlled by RAS/NPS"; Type = "User"},
        @{Name = "Users with a manager"; Type = "User"},
        @{Name = "Users with locked out accounts"; Type = "User"},
        @{Name = "Users with mailbox deleted item retention defaults"; Type = "User"},
        @{Name = "Users with mailbox in specified mailbox store"; Type = "User"},
        @{Name = "Users with no logon script"; Type = "User"},
        @{Name = "Users with no manager"; Type = "User"},
        @{Name = "Users with password stored using reversible encryption"; Type = "User"},
        @{Name = "Users with passwords that expire in the next 5 days"; Type = "User"},
        @{Name = "Users with passwords that expire in the next X days"; Type = "User"},
        @{Name = "Users with passwords that expired in the last 5 days"; Type = "User"},
        @{Name = "Users with passwords that never expire"; Type = "User"},
        @{Name = "Users with specified description"; Type = "User"},
        @{Name = "Users with specified first name"; Type = "User"},
        @{Name = "Users with specified last name"; Type = "User"},
        @{Name = "Users with specified logon script"; Type = "User"},
        @{Name = "Users with specified primary group"; Type = "User"}
    )

    foreach ($query in $queries) {
        $item = New-Object System.Windows.Forms.ListViewItem($query.Name)
        $item.SubItems.Add($query.Type)
        $listView.Items.Add($item)
    }

    $parent.Controls.Add($listView)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Run Query"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(20, 470)
    $parent.Controls.Add($button)

    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"
    $resultBox.Size = New-Object System.Drawing.Size(1100, 150)
    $resultBox.Location = New-Object System.Drawing.Point(20, 510)
    $parent.Controls.Add($resultBox)

    $button.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedItem = $listView.SelectedItems[0]
            $resultBox.Text = switch ($selectedItem.Text) {
                "All users" { Get-ADUser -Filter * -Properties * | Format-List | Out-String }
                "Deleted users" { "Query for Deleted users" }
                "Disabled users" { "Query for Disabled users" }
                "Enabled users" { "Query for Enabled users" }
                "User with specified Employee ID" { "Query for User with specified Employee ID" }
                "User with specified Exchange alias" { "Query for User with specified Exchange alias" }
                "User with specified GUID" { "Query for User with specified GUID" }
                "User with specified name" { "Query for User with specified name" }
                "User with specified primary SMTP address" { "Query for User with specified primary SMTP address" }
                "User with specified SID" { "Query for User with specified SID" }
                "User with specified SID in SID History" { "Query for User with specified SID in SID History" }
                "User with specified SMTP address" { "Query for User with specified SMTP address" }
                "Users created in the last 30 days" { "Query for Users created in the last 30 days" }
                "Users created in the last X days" { "Query for Users created in the last X days" }
                "Users deleted in the last 30 days" { "Query for Users deleted in the last 30 days" }
                "Users in specified company" { "Query for Users in specified company" }
                "Users in specified department" { "Query for Users in specified department" }
                "Users modified in the last 30 days" { "Query for Users modified in the last 30 days" }
                "Users modified in the last X days" { "Query for Users modified in the last X days" }
                "Users not using mailbox size limit defaults" { "Query for Users not using mailbox size limit defaults" }
                "Users not using specified group as primary group" { "Query for Users not using specified group as primary group" }
                "Users that are allowed to dial in" { "Query for Users that are allowed to dial in" }
                "Users that are direct members of specified group" { "Query for Users that are direct members of specified group" }
                "Users that are directly or indirectly members of specified group" { "Query for Users that are directly or indirectly members of specified group" }
                "Users that are not allowed to change password" { "Query for Users that are not allowed to change password" }
                "Users that are not direct members of specified group" { "Query for Users that are not direct members of specified group" }
                "Users that are not directly or indirectly members of specified group" { "Query for Users that are not directly or indirectly members of specified group" }
                "Users that are not protected from deletion" { "Query for Users that are not protected from deletion" }
                "Users that are protected from deletion" { "Query for Users that are protected from deletion" }
                "Users that do not have an Outlook 2010 / Lync photo" { "Query for Users that do not have an Outlook 2010 / Lync photo" }
                "Users that do not require a password" { "Query for Users that do not require a password" }
                "Users that last logged on by specified DC" { "Query for Users that last logged on by specified DC" }
                "Users that will expire in the next 30 days" { "Query for Users that will expire in the next 30 days" }
                "Users where dial in permission is controlled by RAS/NPS" { "Query for Users where dial in permission is controlled by RAS/NPS" }
                "Users with a manager" { "Query for Users with a manager" }
                "Users with locked out accounts" { "Query for Users with locked out accounts" }
                "Users with mailbox deleted item retention defaults" { "Query for Users with mailbox deleted item retention defaults" }
                "Users with mailbox in specified mailbox store" { "Query for Users with mailbox in specified mailbox store" }
                "Users with no logon script" { "Query for Users with no logon script" }
                "Users with no manager" { "Query for Users with no manager" }
                "Users with password stored using reversible encryption" { "Query for Users with password stored using reversible encryption" }
                "Users with passwords that expire in the next 5 days" { "Query for Users with passwords that expire in the next 5 days" }
                "Users with passwords that expire in the next X days" { "Query for Users with passwords that expire in the next X days" }
                "Users with passwords that expired in the last 5 days" { "Query for Users with passwords that expired in the last 5 days" }
                "Users with passwords that never expire" { "Query for Users with passwords that never expire" }
                "Users with specified description" { "Query for Users with specified description" }
                "Users with specified first name" { "Query for Users with specified first name" }
                "Users with specified last name" { "Query for Users with specified last name" }
                "Users with specified logon script" { "Query for Users with specified logon script" }
                "Users with specified primary group" { "Query for Users with specified primary group" }
                default { "No query selected." }
            }
        }
    })
}

# Function to add controls to the Search tab
function Add-SearchControls {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$parent
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Search Active Directory:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $parent.Controls.Add($label)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Size = New-Object System.Drawing.Size(400, 20)
    $searchBox.Location = New-Object System.Drawing.Point(150, 20)
    $parent.Controls.Add($searchBox)

    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search"
    $searchButton.Size = New-Object System.Drawing.Size(100, 30)
    $searchButton.Location = New-Object System.Drawing.Point(560, 20)
    $parent.Controls.Add($searchButton)

    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"
    $resultBox.Size = New-Object System.Drawing.Size(1100, 600)
    $resultBox.Location = New-Object System.Drawing.Point(20, 60)
    $parent.Controls.Add($resultBox)

    $searchButton.Add_Click({
        $resultBox.Text = Get-ADObject -Filter { Name -like "*$($searchBox.Text)*" } -Properties * | Format-List | Out-String
    })
}

# Function to add controls to the Processes/Services tab
function Add-ProcessesServicesControls {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$parent
    )

    # Processes Section
    $processesLabel = New-Object System.Windows.Forms.Label
    $processesLabel.Text = "Processes:"
    $processesLabel.AutoSize = $true
    $processesLabel.Location = New-Object System.Drawing.Point(20, 20)
    $parent.Controls.Add($processesLabel)

    # Ensure you have the necessary assembly loaded
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Create and configure the ListView
    $processesListView = New-Object System.Windows.Forms.ListView
    $processesListView.Size = New-Object System.Drawing.Size(1100, 200)
    $processesListView.Location = New-Object System.Drawing.Point(20, 50)
    $processesListView.View = [System.Windows.Forms.View]::Details
    $processesListView.FullRowSelect = $true
    $processesListView.Columns.Add("Process Name", 400)
    $processesListView.Columns.Add("Process ID", 200)
    $processesListView.Columns.Add("Memory Usage (MB)", 200)
    $processesListView.Columns.Add("CPU Usage (%)", 200)
    $parent.Controls.Add($processesListView)

    # Populate the ListView with data
    $processes = Get-Process
    foreach ($process in $processes) {
        $item = New-Object System.Windows.Forms.ListViewItem($process.Name)
        $item.SubItems.Add($process.Id.ToString())
        $item.SubItems.Add(([math]::Round($process.WorkingSet64 / 1MB, 2)).ToString())
        $item.SubItems.Add($process.CPU.ToString())
        $processesListView.Items.Add($item)
    }

    $refreshProcessesButton = New-Object System.Windows.Forms.Button
    $refreshProcessesButton.Text = "Refresh"
    $refreshProcessesButton.Size = New-Object System.Drawing.Size(100, 30)
    $refreshProcessesButton.Location = New-Object System.Drawing.Point(20, 270)
    $parent.Controls.Add($refreshProcessesButton)

    $killProcessButton = New-Object System.Windows.Forms.Button
    $killProcessButton.Text = "Kill Process"
    $killProcessButton.Size = New-Object System.Drawing.Size(100, 30)
    $killProcessButton.Location = New-Object System.Drawing.Point(130, 270)
    $parent.Controls.Add($killProcessButton)

    $refreshProcessesButton.Add_Click({
        $processesListView.Items.Clear()
        Get-Process | ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem($_.ProcessName)
            $item.SubItems.Add($_.Id)
            $item.SubItems.Add("{0:N2}" -f ($_.WorkingSet64 / 1MB))
            $item.SubItems.Add("{0:N2}" -f ($_.CPU))
            $processesListView.Items.Add($item)
        }
    })

    $killProcessButton.Add_Click({
        if ($processesListView.SelectedItems.Count -gt 0) {
            $processId = $processesListView.SelectedItems[0].SubItems[1].Text
            Stop-Process -Id $processId -Force
            $refreshProcessesButton.PerformClick()
        }
    })

    # Services Section
    $servicesLabel = New-Object System.Windows.Forms.Label
    $servicesLabel.Text = "Services:"
    $servicesLabel.AutoSize = $true
    $servicesLabel.Location = New-Object System.Drawing.Point(20, 320)
    $parent.Controls.Add($servicesLabel)

    $servicesListView = New-Object System.Windows.Forms.ListView
    $servicesListView.Size = New-Object System.Drawing.Size(1100, 200)
    $servicesListView.Location = New-Object System.Drawing.Point(20, 350)
    $servicesListView.View = [System.Windows.Forms.View]::Details
    $servicesListView.FullRowSelect = $true
    $servicesListView.Columns.Add("Service Name", 400)
    $servicesListView.Columns.Add("Display Name", 400)
    $servicesListView.Columns.Add("Status", 150)
    $servicesListView.Columns.Add("Startup Type", 150)
    $parent.Controls.Add($servicesListView)

    $refreshServicesButton = New-Object System.Windows.Forms.Button
    $refreshServicesButton.Text = "Refresh"
    $refreshServicesButton.Size = New-Object System.Drawing.Size(100, 30)
    $refreshServicesButton.Location = New-Object System.Drawing.Point(20, 570)
    $parent.Controls.Add($refreshServicesButton)

    $stopServiceButton = New-Object System.Windows.Forms.Button
    $stopServiceButton.Text = "Stop Service"
    $stopServiceButton.Size = New-Object System.Drawing.Size(100, 30)
    $stopServiceButton.Location = New-Object System.Drawing.Point(130, 570)
    $parent.Controls.Add($stopServiceButton)

    $restartServiceButton = New-Object System.Windows.Forms.Button
    $restartServiceButton.Text = "Restart Service"
    $restartServiceButton.Size = New-Object System.Drawing.Size(100, 30)
    $restartServiceButton.Location = New-Object System.Drawing.Point(240, 570)
    $parent.Controls.Add($restartServiceButton)

    $changeStartupTypeLabel = New-Object System.Windows.Forms.Label
    $changeStartupTypeLabel.Text = "Change Startup Type:"
    $changeStartupTypeLabel.AutoSize = $true
    $changeStartupTypeLabel.Location = New-Object System.Drawing.Point(350, 575)
    $parent.Controls.Add($changeStartupTypeLabel)

    $startupTypeComboBox = New-Object System.Windows.Forms.ComboBox
    $startupTypeComboBox.Size = New-Object System.Drawing.Size(200, 20)
    $startupTypeComboBox.Location = New-Object System.Drawing.Point(470, 575)
    $startupTypeComboBox.Items.AddRange(@("Automatic", "Manual", "Disabled"))
    $parent.Controls.Add($startupTypeComboBox)

    $changeStartupTypeButton = New-Object System.Windows.Forms.Button
    $changeStartupTypeButton.Text = "Change"
    $changeStartupTypeButton.Size = New-Object System.Drawing.Size(100, 30)
    $changeStartupTypeButton.Location = New-Object System.Drawing.Point(680, 570)
    $parent.Controls.Add($changeStartupTypeButton)

    $refreshServicesButton.Add_Click({
        $servicesListView.Items.Clear()
        Get-Service | ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem($_.Name)
            $item.SubItems.Add($_.DisplayName)
            $item.SubItems.Add($_.Status.ToString())
            $item.SubItems.Add($_.StartType.ToString())
            $servicesListView.Items.Add($item)
        }
    })

    $stopServiceButton.Add_Click({
        if ($servicesListView.SelectedItems.Count -gt 0) {
            $serviceName = $servicesListView.SelectedItems[0].SubItems[0].Text
            Stop-Service -Name $serviceName -Force
            $refreshServicesButton.PerformClick()
        }
    })

    $restartServiceButton.Add_Click({
        if ($servicesListView.SelectedItems.Count -gt 0) {
            $serviceName = $servicesListView.SelectedItems[0].SubItems[0].Text
            Restart-Service -Name $serviceName
            $refreshServicesButton.PerformClick()
        }
    })

    $changeStartupTypeButton.Add_Click({
        if ($servicesListView.SelectedItems.Count -gt 0 -and $null -ne $startupTypeComboBox.SelectedItem) {
            $serviceName = $servicesListView.SelectedItems[0].SubItems[0].Text
            $startupType = $startupTypeComboBox.SelectedItem
            Set-Service -Name $serviceName -StartupType $startupType
            $refreshServicesButton.PerformClick()
        }
    })
}

# Function to show the new query form
function Show-NewQueryForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "New Query"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterParent"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter Query:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size(340, 20)
    $textBox.Location = New-Object System.Drawing.Point(20, 50)
    $form.Controls.Add($textBox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Run Query"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(20, 80)
    $button.Add_Click({
        $queryResult = Get-ADObject -LDAPFilter $textBox.Text -Properties *
        [System.Windows.Forms.MessageBox]::Show(($queryResult | Format-Table | Out-String))
    })
    $form.Controls.Add($button)

    [void] $form.ShowDialog()
}

# Function to show the network information form
function Show-NetworkInfoForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Network Information"
    $form.Size = New-Object System.Drawing.Size(600, 400)
    $form.StartPosition = "CenterParent"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Network Information:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.Size = New-Object System.Drawing.Size(540, 300)
    $textBox.Location = New-Object System.Drawing.Point(20, 50)
    $form.Controls.Add($textBox)

    $networkInfo = Get-NetIPAddress | Format-List | Out-String
    $textBox.Text = $networkInfo

    [void] $form.ShowDialog()
}

# Function to show the domain settings form
function Show-DomainSettingsForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Domain Settings"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterParent"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter Domain Name:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size(340, 20)
    $textBox.Location = New-Object System.Drawing.Point(20, 50)
    $form.Controls.Add($textBox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Save"
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(20, 80)
    $button.Add_Click({
        $global:domainName = $textBox.Text
        $form.Close()
    })
    $form.Controls.Add($button)

    [void] $form.ShowDialog()
}

# Function to show the Azure Storage form
function Show-AzureStorageForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Azure Storage"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterParent"

    $accountNameLabel = New-Object System.Windows.Forms.Label
    $accountNameLabel.Text = "Account Name:"
    $accountNameLabel.AutoSize = $true
    $accountNameLabel.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($accountNameLabel)

    $accountNameBox = New-Object System.Windows.Forms.TextBox
    $accountNameBox.Size = New-Object System.Drawing.Size(340, 20)
    $accountNameBox.Location = New-Object System.Drawing.Point(20, 50)
    $form.Controls.Add($accountNameBox)

    $accountKeyLabel = New-Object System.Windows.Forms.Label
    $accountKeyLabel.Text = "Account Key:"
    $accountKeyLabel.AutoSize = $true
    $accountKeyLabel.Location = New-Object System.Drawing.Point(20, 80)
    $form.Controls.Add($accountKeyLabel)

    $accountKeyBox = New-Object System.Windows.Forms.TextBox
    $accountKeyBox.Size = New-Object System.Drawing.Size(340, 20)
    $accountKeyBox.Location = New-Object System.Drawing.Point(20, 110)
    $form.Controls.Add($accountKeyBox)

    $fileURLLabel = New-Object System.Windows.Forms.Label
    $fileURLLabel.Text = "File URL:"
    $fileURLLabel.AutoSize = $true
    $fileURLLabel.Location = New-Object System.Drawing.Point(20, 140)
    $form.Controls.Add($fileURLLabel)

    $fileURLBox = New-Object System.Windows.Forms.TextBox
    $fileURLBox.Size = New-Object System.Drawing.Size(340, 20)
    $fileURLBox.Location = New-Object System.Drawing.Point(20, 170)
    $form.Controls.Add($fileURLBox)

    $containerNameLabel = New-Object System.Windows.Forms.Label
    $containerNameLabel.Text = "Container Name:"
    $containerNameLabel.AutoSize = $true
    $containerNameLabel.Location = New-Object System.Drawing.Point(20, 200)
    $form.Controls.Add($containerNameLabel)

    $containerNameBox = New-Object System.Windows.Forms.TextBox
    $containerNameBox.Size = New-Object System.Drawing.Size(340, 20)
    $containerNameBox.Location = New-Object System.Drawing.Point(20, 230)
    $form.Controls.Add($containerNameBox)

    $blobNameLabel = New-Object System.Windows.Forms.Label
    $blobNameLabel.Text = "Blob Name:"
    $blobNameLabel.AutoSize = $true
    $blobNameLabel.Location = New-Object System.Drawing.Point(20, 260)
    $form.Controls.Add($blobNameLabel)

    $blobNameBox = New-Object System.Windows.Forms.TextBox
    $blobNameBox.Size = New-Object System.Drawing.Size(340, 20)
    $blobNameBox.Location = New-Object System.Drawing.Point(20, 290)
    $form.Controls.Add($blobNameBox)

    $uploadButton = New-Object System.Windows.Forms.Button
    $uploadButton.Text = "Upload"
    $uploadButton.Size = New-Object System.Drawing.Size(100, 30)
    $uploadButton.Location = New-Object System.Drawing.Point(20, 320)
    $uploadButton.Add_Click({
        $accountName = $accountNameBox.Text
        $accountKey = $accountKeyBox.Text
        $fileURL = $fileURLBox.Text
        $containerName = $containerNameBox.Text
        $blobName = $blobNameBox.Text

        $account = [Microsoft.WindowsAzure.Storage.CloudStorageAccount]::Parse("DefaultEndpointsProtocol=https;AccountName=$accountName;AccountKey=$accountKey")
        $blobClient = $account.CreateCloudBlobClient()
        $container = $blobClient.GetContainerReference($containerName)
        $blob = $container.GetBlockBlobReference($blobName)

        $blob.UploadFromFile($fileURL)
        [System.Windows.Forms.MessageBox]::Show("File uploaded successfully.")
    })
    $form.Controls.Add($uploadButton)

    [void] $form.ShowDialog()
}

# Function to show the CrowdStrike form
function Show-CrowdStrikeForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "CrowdStrike"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterParent"

    $clientIDLabel = New-Object System.Windows.Forms.Label
    $clientIDLabel.Text = "Client ID:"
    $clientIDLabel.AutoSize = $true
    $clientIDLabel.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($clientIDLabel)

    $clientIDBox = New-Object System.Windows.Forms.TextBox
    $clientIDBox.Size = New-Object System.Drawing.Size(340, 20)
    $clientIDBox.Location = New-Object System.Drawing.Point(20, 50)
    $form.Controls.Add($clientIDBox)

    $clientSecretLabel = New-Object System.Windows.Forms.Label
    $clientSecretLabel.Text = "Client Secret:"
    $clientSecretLabel.AutoSize = $true
    $clientSecretLabel.Location = New-Object System.Drawing.Point(20, 80)
    $form.Controls.Add($clientSecretLabel)

    $clientSecretBox = New-Object System.Windows.Forms.TextBox
    $clientSecretBox.Size = New-Object System.Drawing.Size(340, 20)
    $clientSecretBox.Location = New-Object System.Drawing.Point(20, 110)
    $form.Controls.Add($clientSecretBox)

    $getDeviceCountButton = New-Object System.Windows.Forms.Button
    $getDeviceCountButton.Text = "Get Device Count"
    $getDeviceCountButton.Size = New-Object System.Drawing.Size(100, 30)
    $getDeviceCountButton.Location = New-Object System.Drawing.Point(20, 140)
    $getDeviceCountButton.Add_Click({
        $clientID = $clientIDBox.Text
        $clientSecret = $clientSecretBox.Text
        $tokenURL = "https://api.crowdstrike.com/oauth2/token"
        $deviceCountURL = "https://api.crowdstrike.com/devices/queries/devices/v1"

        $body = @{
            client_id = $clientID
            client_secret = $clientSecret
            grant_type = "client_credentials"
        }

        $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenURL -Body $body -ContentType "application/x-www-form-urlencoded"
        $accessToken = $tokenResponse.access_token

        $headers = @{
            Authorization = "Bearer $accessToken"
        }

        $deviceCountResponse = Invoke-RestMethod -Method Get -Uri $deviceCountURL -Headers $headers
        $deviceCount = $deviceCountResponse.meta.pagination.total

        [System.Windows.Forms.MessageBox]::Show("Total Devices: $deviceCount")
    })
    $form.Controls.Add($getDeviceCountButton)

    [void] $form.ShowDialog()
}

# Start the GUI
Create-GUI
