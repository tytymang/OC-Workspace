Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar

$start = Get-Date
$end = $start.AddHours(2)

$filter = "[Start] >= '$($start.ToString("yyyy-MM-dd HH:mm"))' AND [Start] <= '$($end.ToString("yyyy-MM-dd HH:mm"))'"
$items = $calendar.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$restricted = $items.Restrict($filter)

foreach($item in $restricted) {
    [PSCustomObject]@{
        Start = $item.Start
        Subject = $item.Subject
    } | ConvertTo-Json -Compress
}
