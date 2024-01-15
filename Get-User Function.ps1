function Get-User {
    # Prompt for a username
    $username = Read-Host "Enter a username"

    # Get the AD user object
    $user = Get-ADUser -Identity $username -Properties "msDS-User-Account-Control-Computed", "PasswordLastSet", "PasswordNeverExpires"

    if (-not $user) {
        Write-Host "User not found"
        return
    }

    # Check if the account is locked out
    if (($user."msDS-User-Account-Control-Computed" -band 0x10) -eq 0x10) {
        $lockStatus = "Locked"
    } else {
        $lockStatus = "Unlocked"
    }

    # Convert the date to UK format for Password Last Set
    $passwordLastSetUK = $user.PasswordLastSet.ToString("dd/MM/yyyy HH:mm:ss")

    # Get password expiry information
    if ($user.PasswordNeverExpires) {
        $passwordExpiryUK = "Never"
    } else {
        $domainPolicy = (Get-ADDefaultDomainPasswordPolicy)
        $passwordExpiryDate = $user.PasswordLastSet + $domainPolicy.MaxPasswordAge
        $passwordExpiryUK = $passwordExpiryDate.ToString("dd/MM/yyyy HH:mm:ss")
    }

    # Display the information
    Write-Host "---------------------------" -ForegroundColor Cyan
    Write-Host "User Details for: $username" -ForegroundColor Green
    Write-Host "---------------------------" -ForegroundColor Cyan
    Write-Host "Status              : $lockStatus"
    Write-Host "Password Last Set   : $passwordLastSetUK"
    Write-Host "Password Expiry     : $passwordExpiryUK"
    Write-Host "---------------------------" -ForegroundColor Cyan

    # If the account is locked, offer to unlock it
    if ($lockStatus -eq "Locked") {
        $unlockChoice = Read-Host "Account is locked. Do you want to unlock it? (Y/N)"
        if ($unlockChoice -eq "Y") {
            Unlock-ADAccount -Identity $username
            Write-Host "Account unlocked." -ForegroundColor Green
        }
    }
}
