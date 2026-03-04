
$Outlook = New-Object -ComObject Outlook.Application
$Calendar = $Outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$Items = $Calendar.Items

# 4월~6월 사이의 [출근] 일정 찾기
$startDate = Get-Date "2026-04-01"
$endDate = Get-Date "2026-06-30"

$targetItems = $Items | Where-Object { 
    $_.Start -ge $startDate -and $_.Start -le $endDate -and $_.Subject -like "*출근*"
}

foreach ($item in $targetItems) {
    $item.ReminderSet = $true
    # 17시간 전 = 17 * 60분 = 1,020분
    $item.ReminderMinutesBeforeStart = 1020 
    $item.Save()
    Write-Host "Updated Reminder: $($item.Start.ToString('yyyy-MM-dd')) - 1020 mins before"
}
