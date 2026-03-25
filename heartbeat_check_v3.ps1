$ErrorActionPreference = "Stop"
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$now = Get-Date
$twoHoursLater = $now.AddHours(2)
$results = @{
    NewVIPEmails = @()
    NextEvent = $null
}

$vips = @("이정훈", "이정우", "이상무", "이영주")
# Last report was ~08:10. Check for emails after 08:10.
$lastReportTime = [DateTime]"2026-03-25 08:10:00"

function Scan-Folders($folder) {
    try {
        # Check for unread VIP emails sent since 08:10
        $filter = "[UnRead] = true AND [SentOn] > '$($lastReportTime.ToString("g"))'"
        $items = $folder.Items.Restrict($filter)
        foreach ($item in $items) {
            foreach ($vip in $vips) {
                if ($item.SenderName -match $vip) {
                    $results.NewVIPEmails += "[$($vip)] $($item.Subject) ($($item.SentOn.ToString('HH:mm')))"
                }
            }
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
$items = $calendar.Items.Restrict($filter) | Sort-Object Start
if ($items.Count -gt 0) {
    $first = $items[0]
    $diff = New-TimeSpan -Start $now -End $first.Start
    $results.NextEvent = "$([Math]::Round($diff.TotalMinutes))분 뒤: $($first.Subject) ($($first.Start.ToString('HH:mm')))"
}

$results | ConvertTo-Json -Compress
