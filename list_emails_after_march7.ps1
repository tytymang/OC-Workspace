$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6)

$startDate = Get-Date "2026-03-08 00:00:00"
$items = $Inbox.Items
$items.Sort("[ReceivedTime]", $true) # Descending

$results = @()
foreach ($item in $items) {
    if ($item.ReceivedTime -ge $startDate) {
        $results += [PSCustomObject]@{
            ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            SenderName = $item.SenderName
            Subject = $item.Subject
        }
    } else {
        break
    }
}

$results | ConvertTo-Json -Compress