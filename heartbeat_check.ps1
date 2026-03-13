$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    # 이메일 체크
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

    # 일정 체크 (향후 2시간)
    $calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar
    $start = Get-Date
    $end = $start.AddHours(2)
    $filter = "[Start] >= '$($start.ToString("g"))' AND [Start] <= '$($end.ToString("g"))'"
    $appointments = $calendar.Items.Restrict($filter)
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