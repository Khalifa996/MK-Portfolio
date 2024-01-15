$Machine = Read-host "Enter Machine Name"
$appName = "Cyberhaven Windows Sensor"
$credentials = Get-Credential -Message "Enter credentials for remote access"

Invoke-Command -ComputerName $Machine -Credential $credentials -ScriptBlock {
    $appName = $using:appName

    $app = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '$appName%'"

    if ($app) {
        $app | ForEach-Object {
            $result = $_.Uninstall()
            if ($result.ReturnValue -eq 0) {
                Write-Output "Application '$appName' has been successfully uninstalled."
            } else {
                Write-Output "Failed to uninstall '$appName'. Error code: $($result.ReturnValue)"
            }
        }
    } else {
        Write-Output "Application '$appName' was not found."
    }
}

pause
