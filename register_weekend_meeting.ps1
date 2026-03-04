$outlook = New-Object -ComObject Outlook.Application
$meeting = $outlook.CreateItem(1)
$meeting.MeetingStatus = 1
$meeting.Subject = "주말 스탭 임원회의"
$meeting.Start = "2026-02-28 08:00"
$meeting.End = "2026-02-28 14:00"
$meeting.Location = "본관접견실 108"
$meeting.Body = "SSKV Groupware 회의실 예약 연동 일정"
$meeting.Save()
Write-Output "Outlook Meeting Registered: $($meeting.Subject)"
