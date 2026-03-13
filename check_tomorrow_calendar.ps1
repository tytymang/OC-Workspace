
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $calendar = $ns.GetDefaultFolder(9)
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    
    $results = @()
    # 내일(3/14) 일정만 필터링
    $tomorrow = (Get-Date).AddDays(1).Date
    $filter = "[Start] >= '" + $tomorrow.ToString("yyyy-MM-dd 00:00") + "' AND [Start] <= '" + $tomorrow.ToString("yyyy-MM-dd 23:59") + "'"
    $tomorrowItems = $items.Restrict($filter)
    
    foreach ($item in $tomorrowItems) {
        $results += [PSCustomObject]@{
            Start   = $item.Start.ToString("HH:mm")
            End     = $item.End.ToString("HH:mm")
            Subject = $item.Subject
            Location = $item.Location
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output $_.Exception.Message
}
