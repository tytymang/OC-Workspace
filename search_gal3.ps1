$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$GAL = $Namespace.GetGlobalAddressList()
$Entries = $GAL.AddressEntries

$count = 0
foreach ($entry in $Entries) {
    if ($entry.Name -match "육하나") {
        $ExUser = $entry.GetExchangeUser()
        if ($ExUser -ne $null) {
            Write-Output "Name: $($entry.Name)"
            Write-Output "Alias: $($ExUser.Alias)"
            Write-Output "JobTitle: $($ExUser.JobTitle)"
            Write-Output "Department: $($ExUser.Department)"
        } else {
            Write-Output "Name: $($entry.Name) (Not an Exchange user)"
        }
        $count++
    }
}
if ($count -eq 0) {
    Write-Output "No one found matching '육하나'"
}
