$FilePathLocation = Read-Host -Prompt 'Please enter the the path of the fold you wish to export'

If ($FilePathLocation -eq $True) {Write-Host "Path Exists"}

Else {Write-Host "Path not found" | Write-Error}
Get-ChildItem $FilePathLocation -Directory -Recurse | ForEach-Object {
    "{0}`n" -f $_.FullName
    Get-ChildItem $_.FullName |
        Select-Object Name |
        Format-Table |
        Out-String
} | Set-Content 'output.txt'





