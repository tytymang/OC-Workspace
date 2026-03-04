
$Outlook = New-Object -ComObject Outlook.Application
$Calendar = $Outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$Items = $Calendar.Items

# 1. 삭제
$Items | ForEach-Object {
    if ($_.Start -gt (Get-Date "2026-03-03")) {
        if ($_.Subject -like "*출근*" -or $_.Subject -like "*異쒓렐*" -or $_.Subject -like "*?*" -or $_.Subject -like "*[?]*") {
            $_.Delete()
        }
    }
}

# 2. 코드 기반 제목 생성 ([출근] 2분기 토요임원회의)
$codeArr = @(91, 52636, 44540, 93, 32, 50, 48516, 44592, 32, 53664, 50836, 51076, 50896, 54940, 51032)
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
    Write-Host "Re-registered successfully: $Date"
}
