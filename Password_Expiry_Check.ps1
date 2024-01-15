# Import the necessary module
Import-Module ActiveDirectory

# Set the OUs you want to check
$ous = @("OU=LN,OU=Accounts,DC=cooley,DC=com", "OU=BR,OU=Accounts,DC=cooley,DC=com")

# Get the current date
$now = Get-Date

# Get the date 7 days from now
$checkDate = (Get-Date).AddDays(7)

# Initialize empty array to store all users
$users = @()

# Get all users in the specific OUs
foreach ($ou in $ous) {
    $users += Get-ADUser -Filter * -SearchBase $ou -Properties "DisplayName", "PasswordLastSet", "PasswordNeverExpires", "Enabled"
}

# Initialize empty array to store users with expiring passwords
$expiringUsers = @()

foreach ($user in $users) {
    # Skip if the user's account is disabled
    if ($user.Enabled -eq $false) {
        continue
    }

    # Skip if the user's password never expires
    if ($user.PasswordNeverExpires -eq $true) {
        continue
    }

    # Skip if the PasswordLastSet is null
    if ($null -eq $user.PasswordLastSet) {
        continue
    }

     # Calculate when the user's password will expire
     $passwordSetDate = $user.PasswordLastSet
     $passwordExpiryDate = $passwordSetDate.AddDays(120) # Password policy is 120 days
 
     # Check if the password has already expired
     if ($now -gt $passwordExpiryDate) {
         # Check if the password expired more than 30 days ago
         if ($now -gt $passwordExpiryDate.AddDays(30)) {
             continue
         }
         $expiryMessage = "Expired"
     }
     # Check if the password will expire in the next 7 days
     elseif ($passwordExpiryDate -le $checkDate) {
         $daysRemaining = ($passwordExpiryDate - $now).Days
         $expiryMessage = "$daysRemaining days remaining"
     }
     else {
         continue
     }
 
     $user | Add-Member -MemberType NoteProperty -Name "PasswordExpiryDate" -Value $passwordExpiryDate -Force
     $user | Add-Member -MemberType NoteProperty -Name "ExpiryMessage" -Value $expiryMessage -Force
     $expiringUsers += $user
 }


# Output users with expiring passwords along with the password expiry date and expiry message
# Sort the users by the PasswordExpiryDate
$sortedExpiringUsers = $expiringUsers | Sort-Object -Property PasswordExpiryDate

$results = $sortedExpiringUsers | Select-Object DisplayName, PasswordExpiryDate, ExpiryMessage | ConvertTo-Html -Head '<style>
body { font-family: Arial, Helvetica, sans-serif; }
table { border-collapse: collapse; width: 100%; }
td, th { border: 1px solid #ddd; padding: 8px; }
th { padding-top: 12px; padding-bottom: 12px; text-align: left; background-color: #4CAF50; color: white; }
</style>'

# Email configuration
$smtpServer = 'Mail.cooley.com'  # Replace with your SMTP server
$smtpPort = 587  # Replace with your SMTP server port
$emailAddress = 'ExpiringPasswords@cooley.com'  # Replace with the sender email address
$displayName = 'ExpiringPasswords'  # Replace with the sender's name

# Define recipients as an array
$recipients = @('mkhalifa@cooley.com', 'rpowar@cooley.com', 'nrobertson@cooley.com', 'ssidhu@cooley.com', 'magyemangduah@cooley.com')

$subject = 'Users with Passwords Expiring within 7 Days'
$body = $results

# Use .NET class to send the email
$mail = New-Object System.Net.Mail.MailMessage
$mail.From = New-Object System.Net.Mail.MailAddress($emailAddress, $displayName)

# Add each recipient to the mail message
foreach ($recipient in $recipients) {
    $mail.To.Add($recipient)
}

$mail.Subject = $subject
$mail.Body = $body
$mail.IsBodyHtml = $true
$smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
# Uncomment below line and replace with appropriate credentials if required
#$smtp.Credentials = New-Object System.Net.NetworkCredential("username", "password") 
$smtp.EnableSsl = $true
$smtp.Send($mail)

