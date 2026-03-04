
$Outlook = New-Object -ComObject Outlook.Application
$Calendar = $Outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$Items = $Calendar.Items

# 1. 4월 ~ 6월 사이의 토요일 일정 중 '출근' 키워드나 깨진 패턴이 포함된 일정 강제 삭제
$startDate = Get-Date "2026-04-01"
$endDate = Get-Date "2026-06-30"

$trash = $Items | Where-Object { 
    $_.Start -ge $startDate -and $_.Start -le $endDate -and 
    ($_.Subject -like "*출근*" -or $_.Subject -like "*횜의*" -or $_.Subject -like "*?*" -or $_.Subject -like "*異쒓렐*")
}

foreach ($item in $trash) {
    $item.Delete()
    Write-Host "DELETED: $($item.Start.ToString('yyyy-MM-dd'))"
}

# 2. 정확한 제목 생성 ([출근] 2분기 토요임원회의)
# [ : 91, 출 : 52636, 근 : 44540, ] : 93, ' ' : 32, 2 : 50, 분 : 48516, 기 : 44592, ' ' : 32, 토 : 53664, 요 : 50836, 임 : 51076, 원 : 50896, 회 : 54924, 의 : 51032
$codeArr = @(91, 52636, 44540, 93, 32, 50, 48516, 44592, 32, 53664, 50836, 51076, 50896, 54924, 51032)
$Subject = ""
foreach($c in $codeArr) { $Subject += [char]$c }

$Schedules = @("2026-04-11", "2026-04-25", "2026-05-16", "2026-05-30", "2026-06-13", "2026-06-27")

foreach ($Date in $Schedules) {
    $Appointment = $Outlook.CreateItem(1)
    $Appointment.Subject = $Subject
    $Appointment.Start = [DateTime]::ParseExact("$Date 08:00", "yyyy-MM-dd HH:mm", $null)
    $Appointment.End = [DateTime]::ParseExact("$Date 14:00", "yyyy-MM-dd HH:mm", $null)
    $Appointment.ReminderSet = $true
    $Appointment.ReminderMinutesBeforeStart = 60
    $Appointment.BusyStatus = 2
    $Appointment.Save()
    Write-Host "FINAL REGISTERED: $Date"
}
