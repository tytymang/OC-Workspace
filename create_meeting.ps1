
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$outlook = New-Object -ComObject Outlook.Application
$appointment = $outlook.CreateItem(1) # olAppointmentItem

$appointment.Subject = "[회의] 재고상태 구분 신규추가 건 확인 및 운영 적용"
$appointment.Start = "2026-02-26 10:30:00"
$appointment.Duration = 60 # 1시간
$appointment.Location = "대회의실(예약 예정)"
$appointment.Body = @"
김종원 선임님, 김승현 대리님, 이동석 팀장님, 육하나님

어제 공유해주신 '재고상태 구분 신규추가' 건과 관련하여, 금일 오전 확인 및 운영 적용을 위한 미팅을 제안드립니다.

[주요 확인 사항]
- 기존 1월 재고 꼬리표와 차이분석
- 1개월 미만 연령 재고 생성 정상 여부
- 재고 배분 금액 확인 (Detail)
- 기존 등록된 재고 소진계획 유지 여부

문제 없으면 기존 1월 재고 꼬리표에 Update 진행 예정입니다.

감사합니다.
"@

$appointment.MeetingStatus = 1 # olMeeting

# 참석자 추가
$recipients = @("김종원", "김승현", "이동석", "육하나")
foreach ($name in $recipients) {
    $recipient = $appointment.Recipients.Add($name)
    $recipient.Resolve()
}

$appointment.Display() # 우선 화면에 띄워서 확인하실 수 있게 합니다.
# $appointment.Send() # 필요시 바로 발송하려면 주석 해제
