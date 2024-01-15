# Elevation Check Function
function Check-Admin {
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "This script must be run as an administrator!"
        exit
    }
}

# First check if the script is running with administrator rights
Check-Admin

# Load the necessary .NET assembly
Add-Type -AssemblyName PresentationFramework

# Define XAML for the GUI layout
[xml]$XAML = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Application Comparison Tool" Height="480" Width="435">
    <Grid>
        <Label Content="Old Computer" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
        <TextBox Name="txtOldComputer" HorizontalAlignment="Left" Height="23" VerticalAlignment="Top" Width="200" Margin="120,10,0,0"/>
        <Label Content="New Computer" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,50,0,0"/>
        <TextBox Name="txtNewComputer" HorizontalAlignment="Left" Height="23" VerticalAlignment="Top" Width="200" Margin="120,50,0,0"/>
        <Button Name="btnCompare" Content="Compare" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="120,100,0,0"/>
        <Label Name="lblStatus" Content="" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,130,0,0"/>
        <ListView Name="lvOutput" HorizontalAlignment="Left" Height="250" Margin="10,160,0,0" VerticalAlignment="Top" Width="400">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Application" Width="200" DisplayMemberBinding="{Binding Application}" />
                    <GridViewColumn Header="Version" Width="200" DisplayMemberBinding="{Binding Version}" />
                </GridView>
            </ListView.View>
        </ListView>
    </Grid>
</Window>
"@

$reader=(New-Object System.Xml.XmlNodeReader $XAML)
$Window=[Windows.Markup.XamlReader]::Load($reader)
$btnCompare = $Window.FindName('btnCompare')
$txtOldComputer = $Window.FindName('txtOldComputer')
$txtNewComputer = $Window.FindName('txtNewComputer')
$lvOutput = $Window.FindName('lvOutput')
$lblStatus = $Window.FindName('lblStatus')

function Get-InstalledSoftware {
    param (
        [Parameter(Mandatory=$true)]
        [string]$computerName
    )
    
    $query = "SELECT Name, Version FROM Win32_Product"
    $options = New-CimSessionOption -Protocol Dcom
    $session = New-CimSession -ComputerName $computerName -SessionOption $options
    $software = Get-CimInstance -Query $query -CimSession $session | 
        Where-Object { $_.Name -notmatch 'Microsoft Visual C\+\+.*|Microsoft .Net.*|Microsoft Privacy .*|.*C\+\+.*|Microsoft Help .*|Microsoft Analysis.*|Microsoft Silverlight.*' }
    $session | Remove-CimSession
    
    return $software
}

function Compare-Software {
    param (
        [Parameter(Mandatory=$true)]
        [string]$computer1,
        [Parameter(Mandatory=$true)]
        [string]$computer2
    )
    
    $software1 = Get-InstalledSoftware -computerName $computer1 | Sort-Object -Property Name
    $software2 = Get-InstalledSoftware -computerName $computer2 | Sort-Object -Property Name
    
    $missing = Compare-Object $software1 $software2 -Property Name -PassThru | Where-Object { $_.SideIndicator -eq '=<' }
    
    return $missing
}

$btnCompare.Add_Click({
    $btnCompare.IsEnabled = $false
    $lblStatus.Content = "Starting comparison... Please wait."

    $computer1 = $txtOldComputer.Text
    $computer2 = $txtNewComputer.Text


    # Check if computer names are not empty
    if ($computer1 -and $computer2) {
        # Creating a new thread to handle the comparison
        $job = Start-Job -ScriptBlock {
            function Get-InstalledSoftware {
                param (
                    [Parameter(Mandatory=$true)]
                    [string]$computerName
                )
            
                $query = "SELECT Name, Version FROM Win32_Product"
                $options = New-CimSessionOption -Protocol Dcom
                $session = New-CimSession -ComputerName $computerName -SessionOption $options
                $software = Get-CimInstance -Query $query -CimSession $session | 
                    Where-Object { $_.Name -notmatch 'Microsoft Visual C\+\+.*|Microsoft .Net.*|Microsoft Privacy .*|.*C\+\+.*|Microsoft Help .*|Microsoft Analysis.*|Microsoft Silverlight.*' }
                $session | Remove-CimSession
            
                return $software
            }

            function Compare-Software {
                param (
                    [Parameter(Mandatory=$true)]
                    [string]$computer1,
                    [Parameter(Mandatory=$true)]
                    [string]$computer2
                )
            
                $software1 = Get-InstalledSoftware -computerName $computer1 | Sort-Object -Property Name
                $software2 = Get-InstalledSoftware -computerName $computer2 | Sort-Object -Property Name
            
                $missing = Compare-Object $software1 $software2 -Property Name -PassThru | Where-Object { $_.SideIndicator -eq '=<' }
            
                return $missing
            }

            # Call Compare-Software function
            $missingApps = Compare-Software -computer1 $args[0] -computer2 $args[1]

            # Return the result
            return $missingApps
        } -ArgumentList $computer1, $computer2

        # Wait for the job to complete
        while ($job.State -eq 'Running') {
            Start-Sleep -Seconds 1
        }

        # Get the result from the job
        $missingApps = Receive-Job -Job $job

        # Check if missingApps is not null or empty
        if ($missingApps) {
            # Update ListView in UI thread
            $lvOutput.ItemsSource = $missingApps | ForEach-Object {
                [PSCustomObject]@{
                    'Application' = $_.Name
                    'Version' = $_.Version
                }
            }
        }

        # Update Label in UI thread
        $lblStatus.Content = "Comparison completed."
        $btnCompare.IsEnabled = $true
    }
    else {
        $lblStatus.Content = "Please enter valid computer names."
        $btnCompare.IsEnabled = $true
    }
})



$Window.ShowDialog() | Out-Null
