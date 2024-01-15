$Machines = Get-ADComputer -Filter 'Name -like "LN-T*"' | Select-Object -ExpandProperty Name
$result = foreach ($Machine in $Machines){
        try {
            
            Test-Connection -ComputerName $Machine -Count 1 -ErrorAction:STOP | Out-Null

        }
        Catch {
            write-warning "Machine: $machine is not pingable - process next item"
            Continue
        }
        #Variables
        
         $Adapter = (Get-NetAdapter -CimSession $machine -Name "Wi-Fi*" ).DriverVersion
         $userName = (Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $machine -ErrorAction:SilentlyContinue).UserName

        if($Adapter -lt "22.170.0.3"){
         
            
            $Status = Write-Host "$userName $Machine $Adapter"
        }
        Elseif ($Adapter -gt "22.150.0.3")  {
            
            $Status = Write-Host "$userName $Machine $Adapter"
        }
        #Build object
        $Output = [ordered]@{
            user=$userName
            machine=$Machine
            IsUpToDate=$Status
            DriverVersion=$Adapter
        }
        new-object -TypeName psobject -Property $Output
    }  

    
$result | Export-Csv -Path C:\Users\Khalifam\Desktop\NetAdapterV.csv -NoTypeInformation
$Output | Export-Csv -Path C:\Users\Khalifam\Desktop\NetAdapterV2.csv
