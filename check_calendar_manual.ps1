Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9)
$items = $calendar.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$now = Get-Date
$twoHoursLater = $now.AddHours(2)

$results = @()
foreach ($item in $items) {
    if ($item.Start -ge $now -and $item.Start -le $twoHoursLater) {
        $results += @{ Subject = $item.Subject; Start = $item.Start.ToString("HH:mm") }
    }
}
$results | ConvertTo-Json -Compress
