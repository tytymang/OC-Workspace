
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    # 1. 메일 확인 (최근 30분 이내)
    $inbox = $namespace.GetDefaultFolder(6)
    $halfHourAgo = (Get-Date).AddMinutes(-30)
    $filter = "[ReceivedTime] >= '$($halfHourAgo.ToString("g"))'"
    $newMails = $inbox.Items.Restrict($filter)
    $mailReport = ""
    foreach ($m in $newMails) {
        $mailReport += "MAIL: $($m.SenderName) - $($m.Subject)`r`n"
    }
    
    # 2. 일정 확인 (향후 2시간 이내)
    $calendar = $namespace.GetDefaultFolder(9)
    $twoHoursLater = (Get-Date).AddHours(2)
    $calItems = $calendar.Items
    $calItems.Sort("[Start]")
    $calItems.IncludeRecurrences = $true
    $upcomingEvents = ""
    foreach ($item in $calItems) {
        if ($item.Start -ge (Get-Date) -and $item.Start -le $twoHoursLater) {
            $upcomingEvents += "CAL: [$($item.Start.ToString('HH:mm'))] $($item.Subject)`r`n"
        }
    }
    
    $res = $mailReport + $upcomingEvents
    [System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\heartbeat_result.txt", $res, [System.Text.Encoding]::Unicode)
} catch {
    [System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\heartbeat_result.txt", "ERROR: $($_.Exception.Message)", [System.Text.Encoding]::Unicode)
}
