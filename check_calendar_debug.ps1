Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9)

# 오늘 전체 일정을 확인하여 필터 작동 여부 검증
$start = (Get-Date).Date
$end = $start.AddDays(1)

$filter = "[Start] >= '$($start.ToString("yyyy-MM-dd HH:mm"))' AND [Start] <= '$($end.ToString("yyyy-MM-dd HH:mm"))'"
$items = $calendar.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$restricted = $items.Restrict($filter)

foreach($item in $restricted) {
    write-output "$($item.Start) - $($item.Subject)"
}
