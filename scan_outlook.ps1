
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# 1. 오늘 일정 확인
$calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar
$start = Get-Date -Hour 0 -Minute 0 -Second 0
$end = $start.AddDays(1)
$filter = "[Start] >= '$($start.ToString("g"))' AND [End] <= '$($end.ToString("g"))'"
$items = $calendar.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$todayEvents = $items.Restrict($filter)

$eventList = @()
foreach ($event in $todayEvents) {
    $eventList += "[ID: $($event.EntryID.Substring(0,8))] $($event.Start.ToString("HH:mm")) - $($event.Subject)"
}

# 2. 주말 사이 안 읽은 메일 확인 (3/14 00:00 이후)
$inbox = $namespace.GetDefaultFolder(6) # olFolderInbox
$lastFriday = (Get-Date 2026-03-14)
$unreadItems = $inbox.Items.Restrict("[UnRead] = true AND [ReceivedTime] >= '$($lastFriday.ToString("g"))'")
$unreadItems.Sort("[ReceivedTime]", $true)

$mailList = @()
foreach ($mail in $unreadItems) {
    $mailList += "[$($mail.ReceivedTime.ToString("MM/dd HH:mm"))] $($mail.SenderName): $($mail.Subject)"
}

$report = @{
    Events = $eventList
    Mails = $mailList
}

$report | ConvertTo-Json | Out-File -FilePath "scan_result.json" -Encoding Unicode
