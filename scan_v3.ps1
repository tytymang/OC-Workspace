
$basePath = "C:\Users\307984\.openclaw\workspace"
Start-Transcript -Path "$basePath\outlook_log.txt"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    # 일정 가져오기
    $calendar = $namespace.GetDefaultFolder(9)
    $today = Get-Date -Hour 0 -Minute 0 -Second 0
    $filter = "[Start] >= '$($today.ToString("g"))' AND [End] <= '$($today.AddDays(1).ToString("g"))'"
    $todayEvents = $calendar.Items.Restrict($filter)
    $eventsStr = "--- 오늘의 일정 ---`r`n"
    foreach ($e in $todayEvents) {
        $eventsStr += "$($e.Start.ToString("HH:mm")): $($e.Subject)`r`n"
    }
    
    # 최근 메일 5건
    $inbox = $namespace.GetDefaultFolder(6)
    $mails = $inbox.Items
    $mails.Sort("[ReceivedTime]", $true)
    $mailsStr = "--- 최근 메일 ---`r`n"
    for ($i=1; $i -le 5; $i++) {
        $m = $mails.Item($i)
        $mailsStr += "[$($m.ReceivedTime.ToString("MM/dd HH:mm"))] $($m.SenderName): $($m.Subject)`r`n"
    }
    
    ($eventsStr + "`r`n" + $mailsStr) | Out-File -FilePath "$basePath\outlook_report.txt" -Encoding Unicode
} catch {
    $_.Exception | Out-File -FilePath "$basePath\error.txt" -Encoding Unicode
}
Stop-Transcript
