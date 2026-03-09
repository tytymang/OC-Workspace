$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6) # 6 is olFolderInbox

$items = $Inbox.Items
$count = $items.Count
$deletedCount = 0

for ($i = $count; $i -ge 1; $i--) {
    $item = $items.Item($i)
    if ($item.SenderName -match "SSCBPM") {
        $subject = $item.Subject
        $item.Delete()
        Write-Output "Deleted: $subject"
        $deletedCount++
    }
}

Write-Output "Total SSCBPM emails deleted: $deletedCount"
