Clear-Host

<#
.SYNOPSIS
    The script imports user data from a CSV file, filters users who have exceeded their mailbox quota and are based in the London office, gathers credentials for an SMTP server, sends personalized emails to each user over quota providing instructions for reducing mailbox size, and logs the emailed users.
.DESCRIPTION
1.The script first imports a CSV file that contains user data. This file is expected to contain user information related to their mailboxes and is located in the given path.
2.The script then filters the user data to identify users who are at or near their email storage limit (identified by the StorageLimitStatus value of 'IssueWarningQuota') and are located in the 'London' office.
3.The script specifies the SMTP server and port it will use to send emails.
4.It prompts the user to input the username and password for the SMTP server and creates a credential object from this information.
5.The script then sets up an empty array to track which users have been emailed, and specifies the location of a text file where this information will be stored.
6.The script enters a loop, iterating over each user who is over their email quota.
  It extracts the first name of the current user.
  It retrieves the user's email address from Active Directory.
  It creates a personalized email message for the user, detailing that their mailbox is nearing its limit and providing steps to reduce the mailbox size.
  It then sends this email to the user and adds their name to the list of users who have been emailed, as well as appending this information to a text file.
7.After emailing all the users who are over their quota, the script sends an email to itself with the list of users that were emailed. This email contains the names of the users and notes that these names have been added to the text file.
.NOTES
    File Name: Mailbox Quota Limit Email.ps1
    Author: Mohamed Khalifa
    Created: 12-May-2023
#>

# Import user data from CSV file
# (Live) Copy and paste this in the quotations below - \\u1script1\reports\Exchange\MailboxQuotaReport\usermailboxreport.csv
$userData = Import-Csv -Path "\\u1script1\reports\Exchange\MailboxQuotaReport\usermailboxreport.csv"

# Filter users whose StorageLimitStatus is 'IssueWarningQuota' and Office is 'London'
$usersOverQuota = $userData | Where-Object { $_.StorageLimitStatus -eq "IssueWarningQuota" -and $_.Office -eq "London" }

# Email parameters
$SmtpServer = "Mail.cooley.com"
$SmtpPort = 587

# Sender's email
$EmailFrom = "Mkhalifa@Cooley.com"
$smtpSubject = "Mailbox Quota Warning"

# Define your credentials for SMTP server
$username = Read-Host -Prompt 'UserName (Elevated Credentials)'
$password = Read-Host -Prompt 'Password (Elevated Credentials)' -AsSecureString
$Credential = New-Object System.Management.Automation.PSCredential ($username, $password)


# List of users that were emailed
$emailedUsers = @()

# Path to your text file
$filePath = "C:\users\khalifam\Desktop\Mailbox Limit_Emailed users.txt"

# Send email to each user over quota
foreach ($user in $usersOverQuota) {
    # Extract the first name from 'DisplayName'
    $firstName = ($user.DisplayName -split ', ')[1]

    # Get the user's email address from Active Directory
    $adUser = Get-ADUser -Identity $user.SamAccountName -Properties EmailAddress
    $smtpTo = $adUser.EmailAddress

    # Construct personalized email body
    $mailBody = @"
Hi $firstName,

Our Outlook mailbox usage report has recently flagged that your mailbox is reaching its limit and could result in emails not being received. 

Here are the steps to reduce your mailbox size:

1. **Check Mailbox Folder Sizes**:
    a. Open Outlook and click on 'File' in the top-left corner.
    b. Under 'Mailbox Settings', click on 'Manage Mailbox Size'.
    c. This will display a list of your folders and their sizes. Take note of folders that are particularly large.

2. **Delete unnecessary emails**. Start with the largest folders identified in step 1.

3. **Archive old emails**: Store old emails that you don’t need immediate access to in an archive.

4. **Empty your 'Deleted Items' folder**: Right-click on the 'Deleted Items' folder and select 'Empty Folder'.

5. **Filing emails in iManage/Link Folders to iManage**.

6. **Unsubscribing from unwanted emails**: Check your subscriptions and unsubscribe from any newsletters or notifications that you no longer wish to receive.

7. **Searching, finding and deleting old broadcast emails** (e.g. sent to "All Hands").

**Benefits**:

1. **Speed and Efficiency**: Smaller mailboxes mean faster and more efficient email systems. Your email client will run more smoothly, and you will spend less time waiting for messages to load or for searches to complete.
2. **Easier to Find Important Emails**: With a smaller, well-organized mailbox, it's easier to find the emails you need when you need them. No more wading through a sea of old, irrelevant messages!
3. **Better Performance**: A large mailbox can slow down the performance of your email client, making it harder to quickly find and respond to important messages. By keeping your mailbox size in check, you can avoid these performance issues.

If you would like assistance with any of the above, please get in contact with a member of the LN-IS team.
Alternatively, you can refer to the training material provided within Workday Learning - https://www.myworkday.com/cooley/learning

Regards,
LN-IS Team
"@


    # Create mail message
    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.From = $EmailFrom
    $mailMessage.To.Add($smtpTo)
    $mailMessage.Subject = $smtpSubject
    $mailMessage.Body = $mailBody

    # Create SMTP client and send the message
    $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
    $smtpClient.EnableSsl = $true
    $smtpClient.Credentials = $Credential
    $smtpClient.Send($mailMessage)

    # Add user to the list of emailed users
    $emailedUsers += $user.DisplayName

    # Append user's name and current date/time to the text file
    Add-Content -Path $filePath -Value ("{0} was emailed on - {1}" -f $user.DisplayName, (Get-Date -Format "dd-MM-yyyy HH:mm:ss"))
}

# Send an email to yourself with the list of users that were emailed
$mailBody = "The following users were emailed RE: Mailbox Quota reaching limit`n`nThe names have been added to $filePath  `n`n" + ($emailedUsers -join "`n")

$mailMessage = New-Object System.Net.Mail.MailMessage
$mailMessage.From = $EmailFrom
$mailMessage.To.Add($EmailFrom) # Send to yourself
$mailMessage.Subject = "Users Emailed"
$mailMessage.Body = $mailBody

$smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
$smtpClient.EnableSsl = $true
$smtpClient.Credentials = $Credential
$smtpClient.Send($mailMessage)