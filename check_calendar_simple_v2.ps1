
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $calendar = $ns.GetDefaultFolder(9)
    
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    
    $results = @()
    # 최근 5개 일정 가져오기
    for ($i = 1; $i -le 10; $i++) {
        $item = $items.Item($i)
        if ($item.Start -ge (Get-Date).Date) {
            $results += [PSCustomObject]@{
                Start   = $item.Start.ToString("yyyy-MM-dd HH:mm")
                Subject = $item.Subject
            }
        }
        if ($results.Count -ge 5) { break }
    }
    $results | ConvertTo-Json
} catch {
    $_.Exception.Message
}
