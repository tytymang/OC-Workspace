
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Items = $Outlook.GetNamespace('MAPI').GetDefaultFolder(6).Items
$Items.Sort("[ReceivedTime]", $true)

$i = 0
foreach ($Item in $Items) {
    if ($Item.Subject -like "*2026*") {
        if ($i -eq 8) {
            Write-Host "---START---"
            Write-Host "SUBJECT: $($Item.Subject)"
            Write-Host "BODY:"
            Write-Host $Item.Body
            Write-Host "---END---"
            break
        }
        $i++
    }
}
