[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$bytes = [Convert]::FromBase64String("7Jyh7ZWY64KY")
$name = [Text.Encoding]::UTF8.GetString($bytes)
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Recip = $Namespace.CreateRecipient($name)
$Recip.Resolve()
if ($Recip.Resolved) {
    $ExUser = $Recip.AddressEntry.GetExchangeUser()
    Write-Output "Found:"
    Write-Output $ExUser.Alias
    Write-Output $ExUser.JobTitle
    Write-Output $ExUser.Department
} else {
    Write-Output "Not resolved: $name"
}
