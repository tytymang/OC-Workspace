$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    $inbox = $namespace.GetDefaultFolder(6)
    $unreadItems = $inbox.Items.Restrict("[UnRead] = true")
    $newEmails = @()
    foreach ($item in $unreadItems) {
        $newEmails += [PSCustomObject]@{
            Sender = $item.SenderName
            Subject = $item.Subject
            Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        }
    }

    $calendar = $namespace.GetDefaultFolder(9)
    $start = Get-Date
    $end = $start.AddHours(2)
    
    # 더 안정적인 날짜 필터 방식
    $filter = "[Start] >= '" + $start.ToString("yyyy-MM-dd HH:mm") + "' AND [Start] <= '" + $end.ToString("yyyy-MM-dd HH:mm") + "'"
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    $appointments = $items.Restrict($filter)
    
    $upcomingMeetings = @()
    foreach ($appt in $appointments) {
        $upcomingMeetings += [PSCustomObject]@{
            Subject = $appt.Subject
            Start = $appt.Start.ToString("yyyy-MM-dd HH:mm")
            Location = $appt.Location
        }
    }

    @{
        Emails = $newEmails
        Meetings = $upcomingMeetings
    } | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}