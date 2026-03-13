Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9)

$now = [DateTime]::Now
$startStr = $now.ToString("yyyy-MM-dd HH:mm")
$endStr = $now.AddHours(24).ToString("yyyy-MM-dd HH:mm")

$filter = "[Start] >= '$startStr' AND [Start] <= '$endStr'"
$items = $calendar.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$restricted = $items.Restrict($filter)

foreach($item in $restricted) {
    write-output "$($item.Start) - $($item.Subject)"
}
