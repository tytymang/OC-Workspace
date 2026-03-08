param($name)
$ErrorActionPreference = 'Stop'
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Recip = $Namespace.CreateRecipient($name)
$Recip.Resolve()
if ($Recip.Resolved) {
    $ExUser = $Recip.AddressEntry.GetExchangeUser()
    Write-Output "Alias: $($ExUser.Alias)"
    Write-Output "ID: $($ExUser.Alias)" # Often alias is the employee ID in this org (e.g., 308579)
} else {
    Write-Output "Not resolved"
}
