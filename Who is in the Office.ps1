$Machines = (Get-ADComputer -ldapfilter "(&(samaccountname=*)(!userAccountControl:1.2.840.113556.1.4.803:=2))" -SearchBase 'OU=Laptops,OU=Win10Modern,OU=LN,OU=Workstations,DC=cooley,DC=com').dnsHostName
$result = @(
    foreach ($Machine in $Machines){
        try {
            #Test connection, remove -quiet otherwise you wont get infos like IP
            $pingResult = Test-Connection -ComputerName $Machine -Count 1 -ErrorAction:Stop
        }
        catch {
            Write-Warning "Machine: $Machine is not pingable - process next item"
            continue
        }
        if($pingResult.Address -like "10.57.*"){
            #Get-WmiObject –ComputerName $Machine –Class Win32_ComputerSystem | Select-Object Username (WMI IS OUTDATED)
            $userName = (Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Machine -ErrorAction:SilentlyContinue).UserName 
            $inOffice = $true
        }
        else {
            $userName = (Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Machine -ErrorAction:SilentlyContinue).UserName
            $inOffice = $false
        }
        #Build object
        $attributes = [ordered]@{
            User = $userName
            Machine = $Machine
            IsInOffice = $inOffice
            IPAddress = $pingResult.Address
        }
        New-Object -TypeName PSObject -Property $attributes
    }
)
$result | Export-Csv -Path C:\Users\khalifam\Desktop\UsersInOffice.csv