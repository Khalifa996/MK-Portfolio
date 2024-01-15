# Define a list of usernames and the desired password
$users = @("User1", "User2", "User3")
$desiredPassword = ConvertTo-SecureString -String "London123!" -AsPlainText -Force

# Loop through the list of usernames and set their password
foreach ($user in $users) {
    # Check if the user exists
    if (Get-ADUser -Filter {SamAccountName -eq $user}) {
        # Set the new password for the user
        Set-ADAccountPassword -Identity $user -NewPassword $desiredPassword -Reset
        # Optionally, force the user to change their password at next login
        Set-AdUser -Identity $user -ChangePasswordAtLogon $true
        Write-Host "Password for $user has been changed."
    } else {
        Write-Host "User $user does not exist."
    }
}
