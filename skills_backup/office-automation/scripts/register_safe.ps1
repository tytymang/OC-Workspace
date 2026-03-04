
param(
    [string[]]$Dates,
    [int[]]$TitleCodes,
    [string]$StartTime = "08:00",
    [string]$EndTime = "14:00",
    [int]$ReminderMin = 60
)

$Outlook = New-Object -ComObject Outlook.Application
$Subject = ""
foreach($c in $TitleCodes) { $Subject += [char]$c }

foreach ($Date in $Dates) {
    $Appointment = $Outlook.CreateItem(1)
    $Appointment.Subject = $Subject
    $Appointment.Start = [DateTime]::ParseExact("$Date $StartTime", "yyyy-MM-dd HH:mm", $null)
    $Appointment.End = [DateTime]::ParseExact("$Date $EndTime", "yyyy-MM-dd HH:mm", $null)
    $Appointment.ReminderSet = $true
    $Appointment.ReminderMinutesBeforeStart = $ReminderMin
    $Appointment.BusyStatus = 2
    $Appointment.Save()
    Write-Host "Registered: $Date ($Subject)"
}
