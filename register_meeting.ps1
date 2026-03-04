[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$outlook = New-Object -ComObject Outlook.Application
$meeting = $outlook.CreateItem(1) # olAppointmentItem

$meeting.MeetingStatus = 1 # olMeeting (이 설정이 있어야 '모임'이 됩니다)
$meeting.Subject = "SCM 구축 협의"
$meeting.Start = "2026-02-27 15:00"
$meeting.End = "2026-02-27 16:00"
$meeting.Location = "1층 108 회의실 (https://sskv.webex.com/sskv/j.php?MTID=m35ed789b69ac6c98c2496b844b12571f)"
$meeting.Body = @"
5. 협의 내용
- 단계별 SCM 구축 전략
- 1단계 과제 상세 협의 > T3 Smart SCM Platform Migration (기능 개선, 이관 전략 等)
"@

$attendees = @("최현구", "이동석", "김제림", "김종원", "강동민", "육하나", "김승현", "정혁곤")
foreach ($person in $attendees) {
    $recipient = $meeting.Recipients.Add($person)
    $recipient.Type = 1 # olRequired
}

# 모임을 보냅니다 (Send 대신 Display 후 주인님이 확인하시거나, 바로 Save/Send 가능)
$meeting.Save()
# $meeting.Send() # 실제로 메일을 발송하려면 이 주석을 해제해야 합니다. 일단 Save하여 '모임'으로 등록합니다.
Write-Output "Meeting registered successfully."
