Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9)
$items = $calendar.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$now = Get-Date
$todayStart = $now.Date
$todayEnd = $todayStart.AddDays(1)
$filter = "[Start] >= '$($todayStart.ToString("yyyy-MM-dd HH:mm"))' AND [End] <= '$($todayEnd.ToString("yyyy-MM-dd HH:mm"))'"
$todayItems = $items.Restrict($filter)
$results = @()
foreach ($item in $todayItems) {
    if ($item.Start -ge $now -and $item.Start -le $now.AddHours(2)) {
        $results += @{ Subject = $item.Subject; Start = $item.Start.ToString("HH:mm") }
    }
}
$results | ConvertTo-Json -Compress
