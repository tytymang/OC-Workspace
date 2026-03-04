
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application

$Schedules = @(
    "2026-04-11", "2026-04-25",
    "2026-05-16", "2026-05-30",
    "2026-06-13", "2026-06-27"
)

foreach ($Date in $Schedules) {
    $Appointment = $Outlook.CreateItem(1) # olAppointmentItem
    $Appointment.Subject = "[출근] 2분기 토요임원회의"
    $Appointment.Start = [DateTime]::ParseExact("$Date 08:00", "yyyy-MM-dd HH:mm", $null)
    $Appointment.End = [DateTime]::ParseExact("$Date 14:00", "yyyy-MM-dd HH:mm", $null)
    $Appointment.ReminderSet = $true
    $Appointment.ReminderMinutesBeforeStart = 60 # 1시간 전 알림
    $Appointment.BusyStatus = 2 # olBusy
    $Appointment.Save()
    Write-Host "REGISTERED: $Date 08:00 - 14:00"
}
