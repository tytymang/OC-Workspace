
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Calendar = $Namespace.GetDefaultFolder(9)
$Items = $Calendar.Items

# 1. 기존 일정 삭제 (미래 일정 중 출근 관련)
$Items | ForEach-Object {
    if ($_.Start -gt (Get-Date "2026-03-03")) {
        # 제목이 깨졌거나 [출근] 패턴인 것들 모두 삭제
        if ($_.Subject -like "*출근*" -or $_.Subject -like "*異쒓렐*" -or $_.Subject -like "*?*" -or $_.Subject -like "*[?]*") {
            $_.Delete()
        }
    }
}

# 2. 유니코드 코드포인트로 제목 생성
# [출근] 2분기 토요임원회의
$s = [char]0x5B + [char]0x出 + [char]0x勤 + [char]0x5D + [char]0x20 + [char]0x32 + [char]0xBD + [char]0xAC + [char]0xAE + [char]0x20 + [char]0xD1 + [char]0xE4 + [char]0xBD + [char]0xBC + [char]0xEC + [char]0xEB + [char]0xED + [char]0xED + [char]0xED + [char]0xED
# 위 방식 대신 더 안전하게 직접 문자를 코드로 변환한 배열을 사용하겠습니다.
$titleArr = @(91, 52636, 44540, 93, 32, 50, 48516, 44592, 32, 53664, 50836, 51076, 50896, 54940, 51032, 32, 51068, 51221, 32, 50504, 45236)
$Subject = ""
foreach($c in $titleArr) { $Subject += [char]$c }

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
    Write-Host "Registered: $Date ($Subject)"
}
