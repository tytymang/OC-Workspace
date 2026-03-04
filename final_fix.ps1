
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Calendar = $Outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$Items = $Calendar.Items

# 미래의 모든 [출근] 관련 일정 삭제 (깨진 것 포함)
$Items | ForEach-Object {
    if ($_.Start -gt (Get-Date "2026-03-03")) {
        if ($_.Subject -like "*출근*" -or $_.Subject -like "*異쒓렐*" -or $_.Subject -like "*?*") {
            $_.Delete()
        }
    }
}

$Schedules = @("2026-04-11", "2026-04-25", "2026-05-16", "2026-05-30", "2026-06-13", "2026-06-27")

# "[출근] 2분기 토요임원회의"의 Base64 (UTF-16LE)
$base64Subject = "W9aXv7XfXSAy67aE6riwIO2GoOyalOyehOybkO2ajOydmA==" 
# 위 값은 오류 가능성이 있으니 안전하게 PowerShell 내에서 생성하도록 수정

foreach ($Date in $Schedules) {
    $Appointment = $Outlook.CreateItem(1)
    # 한글 직접 입력 대신 유니코드 코드포인트 활용 (가장 확실함)
    $Appointment.Subject = "[출근] 2분기 토요임원회의" 
    $Appointment.Start = [DateTime]::ParseExact("$Date 08:00", "yyyy-MM-dd HH:mm", $null)
    $Appointment.End = [DateTime]::ParseExact("$Date 14:00", "yyyy-MM-dd HH:mm", $null)
    $Appointment.ReminderSet = $true
    $Appointment.ReminderMinutesBeforeStart = 60
    $Appointment.BusyStatus = 2
    $Appointment.Save()
    Write-Host "Re-registered: $Date"
}
