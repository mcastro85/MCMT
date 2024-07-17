Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$global:domainName = "Not Connected"

# Define tab structure
$tabDefinitions = @(
    @{Name="Computers"; Type="Computer"; Queries=@("All computers", "Computer by name", "Disabled computers")}
    @{Name="Users"; Type="User"; Queries=@("All users", "User by name", "Disabled users")}
    @{Name="Groups"; Type="Group"; Queries=@("All groups", "Group by name", "Security groups")}
    @{Name="OUs"; Type="OrganizationalUnit"; Queries=@("All OUs", "OU by name")}
    @{Name="GPOs"; Type="GPO"; Queries=@("All GPOs", "GPO by name")}
    @{Name="Printers"; Type="printQueue"; Queries=@("All printers", "Printer by name")}
)

function Create-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "MCMT"
    $form.Size = New-Object System.Drawing.Size(1200, 800)
    $form.StartPosition = "CenterScreen"

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"

    Create-Tabs $tabControl $tabDefinitions
    Add-SpecialTabs $tabControl

    $topPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $topPanel.Dock = "Top"
    $topPanel.AutoSize = $true

    $buttons = @(
        @{Text="New Query"; Action={Show-NewQueryForm}},
        @{Text="Network"; Action={Show-NetworkInfoForm}},
        @{Text="Export Results"; Action={}},
        @{Text="Generate Command Line"; Action={}},
        @{Text="Domain Settings"; Action={Show-DomainSettingsForm}},
        @{Text="Azure Storage"; Action={Show-AzureStorageForm}},
        @{Text="CrowdStrike"; Action={Show-CrowdStrikeForm}}
    )

    foreach ($btn in $buttons) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $btn.Text
        $button.Add_Click($btn.Action)
        $topPanel.Controls.Add($button)
    }

    $domainLabel = New-Object System.Windows.Forms.Label
    $domainLabel.Text = "Domain: $global:domainName"
    $domainLabel.AutoSize = $true
    $topPanel.Controls.Add($domainLabel)

    $form.Controls.AddRange(@($topPanel, $tabControl))
    return $form
}

function Create-Tabs($tabControl, $definitions) {
    foreach ($def in $definitions) {
        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = $def.Name
        
        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Dock = "Left"
        $listBox.Width = 200
        $listBox.Items.AddRange($def.Queries)
        
        $resultBox = New-Object System.Windows.Forms.TextBox
        $resultBox.Multiline = $true
        $resultBox.ScrollBars = "Vertical"
        $resultBox.Dock = "Fill"
        
        $button = New-Object System.Windows.Forms.Button
        $button.Text = "Run Query"
        $button.Dock = "Bottom"
        $button.Add_Click({ 
            $resultBox.Text = Run-Query $def.Type $listBox.SelectedItem 
        })
        
        $tab.Controls.AddRange(@($listBox, $resultBox, $button))
        $tabControl.TabPages.Add($tab)
    }
}

function Add-SpecialTabs($tabControl) {
    $searchTab = New-Object System.Windows.Forms.TabPage
    $searchTab.Text = "Search"
    Add-SearchControls $searchTab

    $processTab = New-Object System.Windows.Forms.TabPage
    $processTab.Text = "Processes/Services"
    Add-ProcessesServicesControls $processTab

    $tabControl.TabPages.AddRange(@($searchTab, $processTab))
}

function Run-Query($type, $query) {
    switch -Wildcard ($query) {
        "All *" { Get-ADObject -Filter {ObjectClass -eq $type} -Properties * | Format-List | Out-String }
        "* by name" { 
            $name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter name:", "Query by Name")
            Get-ADObject -Filter {ObjectClass -eq $type -and Name -like $name} -Properties * | Format-List | Out-String 
        }
        "Disabled *" { Get-ADObject -Filter {ObjectClass -eq $type -and Enabled -eq $false} -Properties * | Format-List | Out-String }
        "Security groups" { Get-ADGroup -Filter {GroupCategory -eq 'Security'} -Properties * | Format-List | Out-String }
        default { "Query not implemented." }
    }
}

function Add-SearchControls($parent) {
    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = "Fill"
    $layout.RowCount = 3
    $layout.ColumnCount = 2

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search"
    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"

    $layout.Controls.AddRange(@(
        (New-Object System.Windows.Forms.Label -Property @{Text="Search:"}),
        $searchBox,
        $searchButton,
        (New-Object System.Windows.Forms.Label -Property @{Text="Results:"}),
        $resultBox
    ))

    $searchButton.Add_Click({
        $resultBox.Text = Get-ADObject -Filter {Name -like "*$($searchBox.Text)*"} -Properties * | Format-List | Out-String
    })

    $parent.Controls.Add($layout)
}

function Add-ProcessesServicesControls($parent) {
    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = "Fill"
    $layout.RowCount = 2
    $layout.ColumnCount = 2

    $processesListView = New-Object System.Windows.Forms.ListView
    $processesListView.View = [System.Windows.Forms.View]::Details
    $processesListView.Columns.AddRange(@("Name", "ID", "Memory (MB)", "CPU (%)"))

    $servicesListView = New-Object System.Windows.Forms.ListView
    $servicesListView.View = [System.Windows.Forms.View]::Details
    $servicesListView.Columns.AddRange(@("Name", "Status", "Start Type"))

    $refreshProcessesButton = New-Object System.Windows.Forms.Button
    $refreshProcessesButton.Text = "Refresh Processes"
    $refreshServicesButton = New-Object System.Windows.Forms.Button
    $refreshServicesButton.Text = "Refresh Services"

    $layout.Controls.AddRange(@(
        (New-Object System.Windows.Forms.Label -Property @{Text="Processes:"}),
        (New-Object System.Windows.Forms.Label -Property @{Text="Services:"}),
        $processesListView,
        $servicesListView,
        $refreshProcessesButton,
        $refreshServicesButton
    ))

    $refreshProcessesButton.Add_Click({
        $processesListView.Items.Clear()
        Get-Process | ForEach-Object {
            $processesListView.Items.Add([PSCustomObject]@{
                Name = $_.Name
                ID = $_.Id
                "Memory (MB)" = [math]::Round($_.WorkingSet64 / 1MB, 2)
                "CPU (%)" = $_.CPU
            })
        }
    })

    $refreshServicesButton.Add_Click({
        $servicesListView.Items.Clear()
        Get-Service | ForEach-Object {
            $servicesListView.Items.Add([PSCustomObject]@{
                Name = $_.Name
                Status = $_.Status
                "Start Type" = $_.StartType
            })
        }
    })

    $parent.Controls.Add($layout)
}

function Show-NewQueryForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "New Query"
    $form.Size = New-Object System.Drawing.Size(400, 200)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = "Fill"
    $layout.RowCount = 3
    $layout.ColumnCount = 2

    $queryBox = New-Object System.Windows.Forms.TextBox
    $resultBox = New-Object System.Windows.Forms.TextBox
    $resultBox.Multiline = $true
    $resultBox.ScrollBars = "Vertical"

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = "Run Query"
    $runButton.Add_Click({
        $resultBox.Text = Get-ADObject -LDAPFilter $queryBox.Text -Properties * | Format-List | Out-String
    })

    $layout.Controls.AddRange(@(
        (New-Object System.Windows.Forms.Label -Property @{Text="LDAP Query:"}),
        $queryBox,
        $runButton,
        (New-Object System.Windows.Forms.Label -Property @{Text="Results:"}),
        $resultBox
    ))

    $form.Controls.Add($layout)
    $form.ShowDialog()
}

function Show-NetworkInfoForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Network Information"
    $form.Size = New-Object System.Drawing.Size(500, 400)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.Dock = "Fill"
    $textBox.Text = Get-NetIPAddress | Format-List | Out-String

    $form.Controls.Add($textBox)
    $form.ShowDialog()
}

function Show-DomainSettingsForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Domain Settings"
    $form.Size = New-Object System.Drawing.Size(300, 150)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = "Fill"
    $layout.RowCount = 3
    $layout.ColumnCount = 2

    $domainBox = New-Object System.Windows.Forms.TextBox
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Add_Click({
        $global:domainName = $domainBox.Text
        $form.Close()
    })

    $layout.Controls.AddRange(@(
        (New-Object System.Windows.Forms.Label -Property @{Text="Domain Name:"}),
        $domainBox,
        $saveButton
    ))

    $form.Controls.Add($layout)
    $form.ShowDialog()
}

function Show-AzureStorageForm {
    # Implementation similar to Show-DomainSettingsForm
    # Add fields for account name, key, file URL, container name, blob name
    # Add upload button with Azure Storage upload logic
}

function Show-CrowdStrikeForm {
    # Implementation similar to Show-DomainSettingsForm
    # Add fields for client ID and secret
    # Add button to get device count with CrowdStrike API call logic
}

# Create and show the main form
$mainForm = Create-MainForm
[void]$mainForm.ShowDialog()