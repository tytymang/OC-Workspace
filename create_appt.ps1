$outlook = New-Object -ComObject Outlook.Application
$appt = $outlook.CreateItem(1)
$subject = [char]48288 + [char]53944 + [char]45224 + [char]32 + [char]52636 + [char]51109
$appt.Subject = $subject
$appt.Start = [datetime]"2026-03-18 00:00:00"
$appt.End = [datetime]"2026-04-03 00:00:00"
$appt.AllDayEvent = $true
$appt.ReminderSet = $true
$appt.ReminderMinutesBeforeStart = 1740
$appt.Display()

$syncObjects = $outlook.Session.SyncObjects
if ($syncObjects.Count -gt 0) {
    $syncObjects.Item(1).Start()
}
