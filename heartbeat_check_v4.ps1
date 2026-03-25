$ErrorActionPreference = "Stop"
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$now = Get-Date
$results = @{
    NewUnread = @()
    NextEvent = $null
}

$lastReportTime = [DateTime]"2026-03-25 08:10:00"

function Scan-Folders($folder) {
    try {
        $filter = "[UnRead] = true AND [SentOn] > '$($lastReportTime.ToString("g"))'"
        $items = $folder.Items.Restrict($filter)
        foreach ($item in $items) {
            $results.NewUnread += "[$($item.SenderName)] $($item.Subject) ($($item.SentOn.ToString('HH:mm')))"
        }
    } catch {}
    foreach ($sub in $folder.Folders) { Scan-Folders $sub }
}

foreach ($root in $namespace.Folders) {
    Scan-Folders $root
}

# Next Event
$calendar = $namespace.GetDefaultFolder(9)
$filter = "[Start] >= '$($now.ToString("g"))' AND [Start] <= '$($now.AddHours(4).ToString("g"))'"
$items = $calendar.Items.Restrict($filter)
if ($null -ne $items) {
    $sorted = $items | Sort-Object Start
    if ($sorted.Count -gt 0) {
        $first = $sorted[0]
        $diff = New-TimeSpan -Start $now -End $first.Start
        $results.NextEvent = "$([Math]::Round($diff.TotalMinutes))분 뒤: $($first.Subject) ($($first.Start.ToString('HH:mm')))"
    }
}

$results | ConvertTo-Json -Compress
