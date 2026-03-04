
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Items = $Outlook.GetNamespace('MAPI').GetDefaultFolder(6).Items
$Items.Sort("[ReceivedTime]", $true)

$i = 0
foreach ($Item in $Items) {
    if ($Item.SenderName -like "*은옥*") {
        Write-Host "---"
        Write-Host "SENDER: $($Item.SenderName)"
        Write-Host "SUBJECT: $($Item.Subject)"
        Write-Host "BODY: $($Item.Body)"
        $i++
    }
    if ($i -ge 5) { break }
}
