
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
    $items = $calendar.Items
    $items.Sort("[Start]", $true) # 최신순
    
    $results = @()
    # 최근 등록된 일정 10개만 확인
    $count = [Math]::Min($items.Count, 10)
    for($i=1; $i -le $count; $i++) {
        $item = $items.Item($i)
        $results += [PSCustomObject]@{
            Subject = $item.Subject
            Start = $item.Start.ToString("yyyy-MM-dd HH:mm")
        }
    }
    $results | ConvertTo-Json
} catch { $_.Exception.Message }
