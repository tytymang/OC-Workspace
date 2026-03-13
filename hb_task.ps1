$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# 1. Unread emails
$inbox = $namespace.GetDefaultFolder(6)
$items = $inbox.Items
$items.Sort("[ReceivedTime]", $true)
$unread = $items | Where-Object { $_.UnRead -eq $true } | Select-Object -First 5
$emailOutput = @()
foreach ($mail in $unread) {
    $dateStr = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
    $sender = $mail.SenderName
    $subject = $mail.Subject
    $emailOutput += "$dateStr`t$sender`t$subject"
}
[System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\hb_emails.txt", ($emailOutput -join "`r`n"), [System.Text.Encoding]::UTF8)

# 2. Today's Calendar
$calendar = $namespace.GetDefaultFolder(9)
$calItems = $calendar.Items
$calItems.Sort("[Start]")
$calItems.IncludeRecurrences = $true
$todayStart = (Get-Date).Date
$todayEnd = $todayStart.AddDays(1)
$filter = "[Start] >= '$($todayStart.ToString("g"))' AND [Start] < '$($todayEnd.ToString("g"))'"
$todayEvents = $calItems.Restrict($filter)
$calOutput = @()
foreach ($evt in $todayEvents) {
    $startStr = $evt.Start.ToString("HH:mm")
    $endStr = $evt.End.ToString("HH:mm")
    $calOutput += "$startStr - $endStr`t$($evt.Subject)"
}
[System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\hb_cal.txt", ($calOutput -join "`r`n"), [System.Text.Encoding]::UTF8)
