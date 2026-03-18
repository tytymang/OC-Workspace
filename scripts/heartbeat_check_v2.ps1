$ErrorActionPreference = 'Stop'
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
    $now = Get-Date
    $cutoff = $now.AddMinutes(-30)

    # Inbox (6)
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    Write-Output "---MAILS---"
    foreach ($m in $items) {
        if ($m.ReceivedTime -lt $cutoff) { break }
        Write-Output "FROM: $($m.SenderName) | SUBJECT: $($m.Subject) | TIME: $($m.ReceivedTime)"
    }

    # Calendar (9)
    $calendar = $namespace.GetDefaultFolder(9)
    $items = $calendar.Items
    $items.Sort("[Start]", $false)
    $items.IncludeRecurrences = $true
    
    $start_range = $now.ToString("yyyy-MM-dd HH:mm")
    $end_range = $now.AddHours(2).ToString("yyyy-MM-dd HH:mm")
    $filter = "[Start] >= '$start_range' AND [Start] <= '$end_range'"
    $upcoming = $items.Restrict($filter)
    
    Write-Output "---EVENTS---"
    foreach ($e in $upcoming) {
        Write-Output "TIME: $($e.Start.ToString('HH:mm')) | TITLE: $($e.Subject)"
    }
} catch {
    Write-Error $_.Exception.Message
}
