
$Outlook = New-Object -ComObject Outlook.Application
$Calendar = $Outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$Items = $Calendar.Items

$startDate = Get-Date "2026-04-01"
$endDate = Get-Date "2026-06-30"

# 모든 일정을 순회하며 해당 기간 내 출근 관련 일정 찾기
foreach ($item in $Items) {
    if ($item.Start -ge $startDate -and $item.Start -le $endDate) {
        $subj = $item.Subject
        # '출근' 혹은 '임원' 키워드 포함 확인 (인코딩 우회)
        if ($subj.Contains([char]52636) -or $subj.Contains([char]51076)) {
            $item.ReminderSet = $true
            # 17시간 전 = 1020분
            $item.ReminderMinutesBeforeStart = 1020
            $item.Save()
            Write-Host "Updated Reminder: $($item.Start.ToString('yyyy-MM-dd HH:mm')) -> 17h Before"
        }
    }
}
