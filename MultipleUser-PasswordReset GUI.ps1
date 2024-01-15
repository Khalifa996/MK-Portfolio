# Load Windows Forms and drawing libraries
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to Update Passwords
function Update-Passwords {
    $users = $textBox.Text -split ','
    $desiredPassword = ConvertTo-SecureString -String $passwordBox.Text -AsPlainText -Force
    $logTextBox.Clear()

    foreach ($user in $users) {
        $user = $user.Trim()
        if ([string]::IsNullOrWhiteSpace($user)) { continue }

        try {
            $userObj = Get-ADUser -Filter "SamAccountName -eq '$user'" -Properties DisplayName
            if ($userObj) {
                Set-ADAccountPassword -Identity $user -NewPassword $desiredPassword -Reset
                if ($changePasswordCheck.Checked) {
                    Set-AdUser -Identity $user -ChangePasswordAtLogon $true
                }
                $logTextBox.AppendText("Password for $($userObj.DisplayName) has been changed.`n")
            } else {
                $logTextBox.AppendText("User $user does not exist.`n")
            }
        } catch {
            $logTextBox.AppendText("Error processing $user $_`n")
        }
    }
}

# Create a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Multi-User Password Reset Tool'
$form.Size = New-Object System.Drawing.Size(400,450)
$form.StartPosition = 'CenterScreen'

# Add a label for username input
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(380,20)
$label.Text = 'Enter usernames (comma-separated):'
$form.Controls.Add($label)

# Add a text box for username input
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(360,20)
$form.Controls.Add($textBox)

# Add a label for password input
$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Location = New-Object System.Drawing.Point(10,70)
$passwordLabel.Size = New-Object System.Drawing.Size(380,20)
$passwordLabel.Text = 'Enter desired password:'
$form.Controls.Add($passwordLabel)

# Add a text box for password input
$passwordBox = New-Object System.Windows.Forms.TextBox
$passwordBox.Location = New-Object System.Drawing.Point(10,90)
$passwordBox.Size = New-Object System.Drawing.Size(360,20)
$passwordBox.UseSystemPasswordChar = $true
$form.Controls.Add($passwordBox)

# Add a checkbox for change password at next logon
$changePasswordCheck = New-Object System.Windows.Forms.CheckBox
$changePasswordCheck.Location = New-Object System.Drawing.Point(10,120)
$changePasswordCheck.Size = New-Object System.Drawing.Size(360,20)
$changePasswordCheck.Text = 'Prompt user to change password on next logon'
$form.Controls.Add($changePasswordCheck)

# Add a button for updating passwords
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(10,150)
$button.Size = New-Object System.Drawing.Size(110,20)
$button.Text = 'Update Passwords'
$button.Add_Click({ Update-Passwords })
$form.Controls.Add($button)

# Add a log text box for feedback
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10,180)
$logTextBox.Size = New-Object System.Drawing.Size(360,250)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = 'Vertical'
$logTextBox.ReadOnly = $true
$form.Controls.Add($logTextBox)

# Show the form
$form.ShowDialog()
