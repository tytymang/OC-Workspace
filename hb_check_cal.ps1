try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
    $startTime = Get-Date
    $endTime = $startTime.AddHours(2)
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    $filter = "[Start] >= '" + $startTime.ToString("g") + "' AND [Start] <= '" + $endTime.ToString("g") + "'"
    $recentApps = $items.Restrict($filter)

    $results = @()
    foreach ($app in $recentApps) {
        $results += @{ Subject = $app.Subject; Start = $app.Start.ToString("HH:mm") }
    }
    if ($results.Count -eq 0) { "[]" } else { $results | ConvertTo-Json }
} catch { $_.Exception.Message } | Out-File "hb_cal.json" -Encoding UTF8
