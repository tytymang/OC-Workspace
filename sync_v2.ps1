
$Outlook = New-Object -ComObject Outlook.Application
$Calendar = $Outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$Items = $Calendar.Items

$startDate = Get-Date "2026-04-01"
$endDate = Get-Date "2026-06-30"

$count = 0
foreach ($item in $Items) {
    # 날짜 범위 확인 (4월~6월)
    if ($item.Start -ge $startDate -and $item.Start -le $endDate) {
        # '출근' ([char]52636) 키워드 확인
        if ($item.Subject.Contains([char]52636)) {
            $item.ReminderSet = $true
            $item.ReminderMinutesBeforeStart = 1020 # 17시간 전
            $item.Save()
            Write-Host "Sync: $($item.Start.ToString('yyyy-MM-dd HH:mm')) -> 15:00 Reminder OK"
            $count++
        }
    }
}

if ($count -gt 0) {
    # 모든 폴더 동기화 트리거
    $Outlook.Session.SyncObjects.Item(1).Start()
    Write-Host "Google Calendar Synchronization Triggered."
}
