$Machines = Get-Content -Path "C:\Users\Khalifam\Desktop\LNLaptops.txt"

$result = foreach ($Machine in $Machines) 
 { 
    if (Test-Connection -ComputerName $Machine -Count 1 -Quiet) {
        Invoke-Command -ComputerName $Machine -ScriptBlock {
            Get-NetAdapterAdvancedProperty -Name "Wi-Fi*" -DisplayName "Roaming Aggressiveness" | Select-Object Name,DisplayName,DisplayValue,SystemName
   }
  }
 }

 $result | Export-Csv -Path 'C:\Users\Khalifam\Desktop\NetAdapterAggresivenessInfo.csv' -NoTypeInformation
    