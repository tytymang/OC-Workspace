
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $calendar = $ns.GetDefaultFolder(9) # olFolderCalendar
    $targetDate = (Get-Date).Date
    
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    
    # 오늘 일정 검색
    $filter = "[Start] >= '" + $targetDate.ToString("yyyy-MM-dd 00:00") + "' AND [Start] <= '" + $targetDate.ToString("yyyy-MM-dd 23:59") + "'"
    $todayItems = $items.Restrict($filter)
    
    $results = @()
    foreach ($item in $todayItems) {
        $results += [PSCustomObject]@{
            Start   = $item.Start.ToString("HH:mm")
            End     = $item.End.ToString("HH:mm")
            Subject = $item.Subject
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Error $_.Exception.Message
}
