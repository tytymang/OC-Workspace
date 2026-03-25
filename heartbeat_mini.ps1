Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$cal = $ns.GetDefaultFolder(9)
$items = $cal.Items
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
