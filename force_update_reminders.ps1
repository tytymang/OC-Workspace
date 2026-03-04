
$Outlook = New-Object -ComObject Outlook.Application
$Calendar = $Outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$Items = $Calendar.Items

$Dates = @("2026-04-11", "2026-04-25", "2026-05-16", "2026-05-30", "2026-06-13", "2026-06-27")

foreach ($d in $Dates) {
    # 해당 날짜 오전 8시 일정 필터링
    $targetDate = [DateTime]::ParseExact("$d 08:00", "yyyy-MM-dd HH:mm", $null)
    $found = $Items | Where-Object { $_.Start -eq $targetDate }
    
    foreach ($item in $found) {
        $item.ReminderSet = $true
        # 17시간 전 = 1020분
        $item.ReminderMinutesBeforeStart = 1020
        $item.Save()
        Write-Host "Updated: $d 08:00 - Reminder set to 17h (1020m) before."
    }
}
