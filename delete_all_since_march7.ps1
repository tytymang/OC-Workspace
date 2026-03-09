$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6)

$items = $Inbox.Items
$items.Sort("[ReceivedTime]", $true) # Descending (newest first)

$startDate = Get-Date "2026-03-07 00:00:00"

$toDelete = @()

foreach ($item in $items) {
    if ($null -ne $item.ReceivedTime) {
        if ($item.ReceivedTime -ge $startDate) {
            $toDelete += $item
        } else {
            break # Since it's sorted descending, older items follow
        }
    }
}

$deletedCount = 0
foreach ($item in $toDelete) {
    $item.Delete()
    $deletedCount++
}

Write-Output "Total deleted: $deletedCount"
