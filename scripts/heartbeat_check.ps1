$ErrorActionPreference = 'SilentlyContinue'
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$now = Get-Date
$cutoff = $now.AddMinutes(-30)

$inbox = $namespace.GetDefaultFolder(6)
$mails = $inbox.Items | Where-Object { $_.ReceivedTime -gt $cutoff }
Write-Output "---MAILS---"
foreach ($m in $mails) {
    Write-Output "FROM: $($m.SenderName) | SUBJECT: $($m.Subject)"
}

$calendar = $namespace.GetDefaultFolder(9)
$filter = "[Start] >= '$($now.ToString('g'))' AND [Start] <= '$($now.AddHours(2).ToString('g'))'"
$items = $calendar.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$events = $items.Restrict($filter)
Write-Output "---EVENTS---"
foreach ($e in $events) {
    Write-Output "TIME: $($e.Start.ToString('HH:mm')) | TITLE: $($e.Subject)"
}
