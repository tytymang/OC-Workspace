
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Items = $Outlook.GetNamespace('MAPI').GetDefaultFolder(6).Items
$Items.Sort("[ReceivedTime]", $true)

$i = 0
foreach ($Item in $Items) {
    if ($Item.Subject -like "*2026*") {
        Write-Host "ID: $i"
        Write-Host "S: $($Item.SenderName)"
        Write-Host "T: $($Item.Subject)"
        Write-Host "---"
        $i++
    }
    if ($i -ge 20) { break }
}
