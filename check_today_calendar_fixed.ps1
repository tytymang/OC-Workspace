[CmdletBinding()]
Param()
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar
$now = Get-Date
$endOfDay = $now.Date.AddDays(1)
$items = $calendar.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$restriction = "[Start] >= '" + $now.ToString("g") + "' AND [Start] <= '" + $endOfDay.ToString("g") + "'"
$todayItems = $items.Restrict($restriction)

foreach ($item in $todayItems) {
    Write-Output "Subject: $($item.Subject) | Start: $($item.Start) | Location: $($item.Location)"
}
