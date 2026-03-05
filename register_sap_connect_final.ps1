
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $appt = $outlook.CreateItem(1) # olAppointmentItem
    
    # 유니코드 조립 (SAP Connect Day)
    # 제목: [세미나] SAP Connect Day (이정우 부사장 공유)
    $subject = "[SAP Connect Day] " + [char]51060 + [char]51221 + [char]50864 + " " + [char]48512 + [char]49324 + [char]51109 + " " + [char]44277 + [char]50976
    
    $appt.Subject = $subject
    $appt.Location = "Grand InterContinental Seoul Parnas"
    $appt.Body = "2/10 수신된 SAP Connect Day 세미나 초청 건입니다."
    
    # 시간 설정: 2026-03-19 14:00 ~ 17:30
    $appt.Start = Get-Date -Year 2026 -Month 3 -Day 19 -Hour 14 -Minute 0 -Second 0
    $appt.End = Get-Date -Year 2026 -Month 3 -Day 19 -Hour 17 -Minute 30 -Second 0
    
    # 1일 전(1440분) 알림 설정
    $appt.ReminderSet = $true
    $appt.ReminderMinutesBeforeStart = 1440
    
    $appt.Save()
    
    # 구글 캘린더 동기화 강제 트리거 (Send/Receive)
    $outlook.Session.SyncObjects.Item(1).Start()
    
    "SUCCESS"
} catch {
    $_.Exception.Message
}
