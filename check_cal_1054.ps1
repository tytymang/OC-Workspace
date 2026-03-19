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
        $results += @{
            Subject = $app.Subject
            Start = $app.Start.ToString("yyyy-MM-dd HH:mm")
        }
    }
    if ($results.Count -eq 0) {
        "[]" | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\hb_cal_1054.json" -Encoding UTF8
    } else {
        $results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\hb_cal_1054.json" -Encoding UTF8
    }
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\hb_cal_err_1054.txt"
}
