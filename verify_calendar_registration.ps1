
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9) # 9 = olFolderCalendar
    
    # 2026-03-19 일정을 검색
    $start = "2026-03-19 00:00"
    $end = "2026-03-19 23:59"
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    
    $foundItems = @()
    foreach ($item in $items) {
        if ($item.Start -ge [datetime]$start -and $item.Start -le [datetime]$end) {
            $foundItems += [PSCustomObject]@{
                Subject = $item.Subject
                Start = $item.Start.ToString("yyyy-MM-dd HH:mm")
                Location = $item.Location
            }
        }
    }
    
    if ($foundItems.Count -gt 0) {
        $foundItems | ConvertTo-Json
    } else {
        "NOT_FOUND"
    }
} catch {
    $_.Exception.Message
}
