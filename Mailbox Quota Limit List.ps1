Clear-Host

<#
.SYNOPSIS
    The script imports user data from a CSV file and identifies users based in the London office who have exceeded their mailbox quota.

.DESCRIPTION
    The script performs the following tasks:

    1. Imports user mailbox data from a CSV file.
    2. Filters and identifies users from the London office who have exceeded their Issue Warning Quota.
    3. Collates a list of these users, displaying their name and the percentage of their mailbox usage.
    4. This list is saved to a text file on the desktop for easy reference.
    5. A consolidated email containing this list is sent to the administrator for awareness.

.NOTES
    File Name: Mailbox Quota Limit List.ps1
    Author: Mohamed Khalifa
    Created: 21-AUG-2023

.EXAMPLE
    .\Mailbox Quota Limit List.ps1

    Upon execution, the script prompts the administrator for elevated SMTP credentials. After processing the data, the script sends an email to the administrator and saves a report on the desktop.

#>


# Import user data from CSV file
$userData = Import-Csv -Path "\\u1script1\reports\Exchange\MailboxQuotaReport\usermailboxreport.csv"

# Define the path to the report file on your desktop
$reportFilePath = "c:\Users\$env:CooleyPrimaryUserName\Desktop\MailboxUsageReport.txt"

# Ensure any previous report is cleared/overwritten
if (Test-Path $reportFilePath) {
    Remove-Item $reportFilePath
}

$flaggedUsersReport = @()

foreach ($user in $userData) {
    # Filter out only users from the "London" office
    if ($user.Office -eq "London") {
        try {
            # Convert IssueWarningQuotaGB to MB for comparison
            $warningQuotaMB = [double]::Parse($user.IssueWarningQuotaGB) * 1024

            # Ensure the MailboxSizeMB can also be converted
            $mailboxSize = [double]::Parse($user.MailboxSizeMB)

            # Check if MailboxSizeMB exceeds the IssueWarningQuotaGB
            if ($mailboxSize -gt $warningQuotaMB) {
                # Calculate percentage used
                $percentageUsed = ($mailboxSize / $warningQuotaMB) * 100

                # Add user (using DisplayName) and their percentage to the report list
                $reportLine = "{0} has used {1:N2}% of their mailbox warning limit." -f $user.DisplayName, $percentageUsed
                
                $flaggedUsersReport += $reportLine
                
                # Append this line to the report file on your desktop
                Add-Content -Path $reportFilePath -Value $reportLine
            }

            # Check if MailboxSizeMB exceeds the ProhibitSendQuotaGB (in GB)
            $prohibitSendQuotaGB = [double]::Parse($user.ProhibitSendQuotaGB)
            if ($mailboxSize -gt ($prohibitSendQuotaGB * 1024)) {
                # Add a second paragraph for users who can no longer send mail
                $secondParagraph = "{0} can no longer send mail as their mailbox size exceeds the ProhibitSendQuota." -f $user.DisplayName
                $flaggedUsersReport += $secondParagraph
                Add-Content -Path $reportFilePath -Value $secondParagraph
            }
        } catch {
            Write-Host "Error processing user $($user.DisplayName): $_"
        }
    }
}

# Email parameters
$SmtpServer = "Mail.cooley.com"
$SmtpPort = 587

# Sender's email
$EmailFrom = "mkhalifa@Cooley.com"
$smtpSubject = "Mailbox Quota Warning"

# Define your credentials for SMTP server
$username = Read-Host -Prompt 'UserName (Elevated Credentials)'
$password = Read-Host -Prompt 'Password (Elevated Credentials)' -AsSecureString
$Credential = New-Object System.Management.Automation.PSCredential ($username, $password)

# Create the email body
$mailBody = "The following users in London office were flagged due to reaching their Mailbox Quota limit:`n`n" + ($flaggedUsersReport -join "`n") + "`n`nThe report has also been saved to $reportFilePath."

$mailMessage = New-Object System.Net.Mail.MailMessage
$mailMessage.From = $EmailFrom
$mailMessage.To.Add($EmailFrom) # Send to yourself
$mailMessage.Subject = "Users Flagged for Mailbox Quota Limit in London"
$mailMessage.Body = $mailBody

$smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
$smtpClient.EnableSsl = $true
$smtpClient.Credentials = $Credential
$smtpClient.Send($mailMessage)
