
$Outlook = New-Object -ComObject Outlook.Application
$Calendar = $Outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$Items = $Calendar.Items

$startDate = Get-Date "2026-04-01"
$endDate = Get-Date "2026-06-30"

# '[출근]' 키워드가 포함된 일정만 정밀 타격
foreach ($item in $Items) {
    if ($item.Start -ge $startDate -and $item.Start -le $endDate) {
        $subj = $item.Subject
        # '출근' 키워드 확인 (유니코드 52636)
        if ($subj.Contains([char]52636)) {
            $item.ReminderSet = $true
            $item.ReminderMinutesBeforeStart = 1020 # 17시간 전 (금요일 15:00)
            $item.Save()
            Write-Host "Updated Google-Synced Reminder: $($item.Start.ToString('yyyy-MM-dd HH:mm')) -> 17h Before"
        }
    }
}

# Outlook 강제 동기화 트리거 (보내기/받기 유사 효과)
$Outlook.GetNamespace("MAPI").SendAndReceive($true)
Write-Host "Syncing with Google Calendar server..."
