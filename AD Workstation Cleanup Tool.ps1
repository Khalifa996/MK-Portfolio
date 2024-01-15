Clear-Host
#Elevation Check Function
function Check-Admin {
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        # Display a message and exit the script instead of launching a new process
        Write-Host "This script must be run as an administrator!"
        exit
    }
}

# Load Assembly Function
function Load-Assemblies {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
}

# XAML Load and Parse Functi
function Parse-XAML {
    param(
        [Parameter(Mandatory=$true)]
        [string] $xaml
    )

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xaml)
    return [Windows.Markup.XamlReader]::Load($reader)
}

# Load Machines from OU Function
function Load-Machines {
    param(
        [Parameter(Mandatory=$true)]
        [string] $ou,
        [Parameter(Mandatory=$true)]
        [string[]] $machineNames
    )

    $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, "cooley.com", $ou)
    $qbe = New-Object System.DirectoryServices.AccountManagement.ComputerPrincipal $ctx

    $ps = New-Object System.DirectoryServices.AccountManagement.PrincipalSearcher
    $ps.QueryFilter = $qbe

    $regex = $machineNames -join "|"
    $regex = "^CN=(" + $regex + ")"

    return $ps.FindAll() | Where-Object { $_.DistinguishedName -match $regex }
}

# Add Move Click Event Function
function Add-MoveClickEvent {
    param(
        [Parameter(Mandatory=$true)]
        $button,
        [Parameter(Mandatory=$true)]
        $window
    )

    $button.Add_Click({
        $selectedMachine = $window.FindName("MachineList").SelectedItem
        if ($selectedMachine -ne $null) {
            $sourceOU = "OU=Computers - Disabled,DC=cooley,DC=com"
            $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, "cooley.com", $sourceOU)

            # Checking if the machine name starts with BR-T*
            if ($selectedMachine.Name -like "BR-T*") {
                $targetOU = "OU=Laptops,OU=Win10Modern,OU=BR,OU=Workstations,DC=cooley,DC=com"
            } else {
                $targetOU = "OU=Laptops,OU=Win10Modern,OU=LN,OU=Workstations,DC=cooley,DC=com"
            }

            $sourcePath = "LDAP://" + $selectedMachine.DistinguishedName
            $targetPath = "LDAP://" + $targetOU

            $sourceEntry = New-Object System.DirectoryServices.DirectoryEntry $sourcePath
            $targetEntry = New-Object System.DirectoryServices.DirectoryEntry $targetPath

            $sourceEntry.PSBase.MoveTo($targetEntry, "CN=" + $selectedMachine.Name)

            $sourceEntry = [System.DirectoryServices.AccountManagement.ComputerPrincipal]::FindByIdentity($ctx, $selectedMachine.Name)

            # Update description with move date
            $sourceEntry.Description = "Moved back to target OU on $(Get-Date -Format 'dd/MM/yyyy')"
            $sourceEntry.Enabled = $true
            $sourceEntry.Save()
            
            $window.FindName("MachineList").Items.Remove($selectedMachine)
        }
    })
}



# Add MarkChecked Click Event Function
function Add-MarkCheckedClickEvent {
    param(
        [Parameter(Mandatory=$true)]
        $button,
        [Parameter(Mandatory=$true)]
        $window
    )

    $button.Add_Click({
        $selectedMachine = $window.FindName("MachineList").SelectedItem
        if ($selectedMachine -ne $null) {
            $targetPath = "LDAP://" + $selectedMachine.DistinguishedName
            $targetEntry = New-Object System.DirectoryServices.DirectoryEntry $targetPath

            # Get the current date and format it
            $currentDate = Get-Date -Format "dd/MM/yyyy"

            # Update description with the migration note and the current date
            $newDescription = "##Checked## - " + $currentDate
            $targetEntry.InvokeSet("Description", $newDescription)
            $targetEntry.CommitChanges()
        }
    })
}


# Add Copy Click Event Function
function Add-CopyClickEvent {
    param(
        [Parameter(Mandatory=$true)]
        $button,
        [Parameter(Mandatory=$true)]
        $window
    )

    $button.Add_Click({
        $selectedMachine = $window.FindName("MachineList").SelectedItem
        if ($selectedMachine -ne $null) {
            # Copy machine name to clipboard
            $selectedMachine.Name | Set-Clipboard
        }
    })
}

# Add Refresh Click Event Function
function Add-RefreshClickEvent {
    param(
        [Parameter(Mandatory=$true)]
        $button,
        [Parameter(Mandatory=$true)]
        $window,
        [Parameter(Mandatory=$true)]
        [string] $ou,
        [Parameter(Mandatory=$true)]
        [string[]] $machineNames
    )

    $button.Add_Click({
        $machineList = $window.FindName("MachineList")

        # Clear current items
        $machineList.Items.Clear()

        # Re-load machines
        $machines = Load-Machines $ou $machineNames

        foreach ($machine in $machines) {
            $description = $machine.Description -replace '(?m)^[\s-]+', ''

            $computerName = $machine.Name
            $computer = Get-ADComputer -Identity $computerName -Properties LastLogonTimestamp
            $lastLogonTimestamp = $computer.LastLogonTimestamp

            if ($lastLogonTimestamp) {
                $lastLogonTime = [DateTime]::FromFileTime($lastLogonTimestamp)
            } else {
                $lastLogonTime = "Last Logon Timestamp not available."
            }

            $item = New-Object PSObject -Property @{
                Name = $machine.Name
                Description = $description
                LastLogonTimeStamp = $lastLogonTime
                DistinguishedName = $machine.DistinguishedName
            }
            $machineList.Items.Add($item)
        }
    })
}

# Sort Columns Function - NEW!
function Sort-Columns {
    param(
        [Parameter(Mandatory=$true)]
        $sender,
        [Parameter(Mandatory=$true)]
        $e
    )

    $headerClicked = $e.OriginalSource -as [System.Windows.Controls.GridViewColumnHeader]
    if ($headerClicked -ne $null -and $headerClicked.Role -ne 'Padding') {
        $listView = $sender
        $binding = $headerClicked.Column.DisplayMemberBinding.Path.Path
        $listView.Items.SortDescriptions.Clear()
        $sortDescription = New-Object System.ComponentModel.SortDescription $binding, 'Ascending'
        $listView.Items.SortDescriptions.Add($sortDescription)
        $listView.Items.Refresh()
    }
}


# Main Script

# Elevation Check
Check-Admin

# Load required assemblies
Load-Assemblies

# Load XAML

$xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='Workstation Cleanup' Height='400' Width='565'>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row='0' Text='Select a machine to move:' Margin='10'/>
        <ListView Grid.Row='1' Name='MachineList' Margin='10'>
            <ListView.View>
                <GridView>
                    <GridViewColumn Header='Machine Name' DisplayMemberBinding='{Binding Name}'/>
                    <GridViewColumn Header='Description' DisplayMemberBinding='{Binding Description}'/>
                    <GridViewColumn Header='Last Logon Time' DisplayMemberBinding='{Binding LastLogonTimeStamp}'/>
                </GridView>
            </ListView.View>
        </ListView>
        <StackPanel Grid.Row='2' Orientation='Horizontal'>
            <Button Name='MarkChecked' Content='Mark as Checked' Margin='10' HorizontalAlignment='Left'/>
            <Button Name='CopyButton' Content='Copy Selected Machine Name' Margin='10' HorizontalAlignment='Left'/>
            <Button Name='MoveButton' Content='Move/Enable Selected Machine' Margin='10' HorizontalAlignment='Left'/>
            <Button Name='RefreshButton' Content='Refresh' Margin='10' HorizontalAlignment='Right'/>
        </StackPanel>
    </Grid>
</Window>
"@

# Parse XAML
$window = Parse-XAML $xaml

# Load machines from the specified OU
$ou = "OU=Computers - Disabled,DC=cooley,DC=com"
$machineNames = "LN-T14*", "LN-T49*", "LN-T48*", "BR-T1*", "BR-T49*"

$machines = Load-Machines $ou $machineNames
$machineList = $window.FindName("MachineList")
$machineList.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, [System.Windows.RoutedEventHandler]{ Sort-Columns $machineList $_ })


foreach ($machine in $machines) {
    $description = $machine.Description -replace '(?m)^[\s-]+', ''

    $computerName = $machine.Name
    $computer = Get-ADComputer -Identity $computerName -Properties LastLogonTimestamp
    $lastLogonTimestamp = $computer.LastLogonTimestamp

    if ($lastLogonTimestamp) {
        $lastLogonTime = [DateTime]::FromFileTime($lastLogonTimestamp)
    } else {
        $lastLogonTime = "Last Logon Timestamp not available."
    }

    $item = New-Object PSObject -Property @{
        Name = $machine.Name
        Description = $description
        LastLogonTimeStamp = $lastLogonTime
        DistinguishedName = $machine.DistinguishedName
    }
    $machineList.Items.Add($item) | Out-Null
}

# Add Click Events

$markCheckedButton = $window.FindName("MarkChecked")
$copyButton = $window.FindName("CopyButton")
$moveButton = $window.FindName("MoveButton")
$refreshButton = $window.FindName("RefreshButton")

Add-MarkCheckedClickEvent $markCheckedButton $window
Add-CopyClickEvent $copyButton $window
Add-MoveClickEvent $moveButton $window
Add-RefreshClickEvent $refreshButton $window $ou $machineNames

# Show Window
$null = $window.ShowDialog() | Out-Null