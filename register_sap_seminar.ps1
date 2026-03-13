
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $appt = $outlook.CreateItem(1) # olAppointmentItem
    
    # [char] 코드를 이용한 유니코드 조립 (SAP Build 통합 세미나)
    # S:83, A:65, P:80,  :32, B:66, u:117, i:105, l:108, d:100,  :32, 통:53685, 합:54633,  :32, 세:49464, 미:48120, 나:45208
    $subject = "SAP Build " + [char]53685 + [char]54633 + [char]32 + [char]49464 + [char]48120 + [char]45208 + " (" + [char]51060 + [char]51221 + [char]50864 + [char]32 + [char]48512 + [char]49324 + [char]51109 + [char]45784 + ")"
    
    $appt.Subject = $subject
    $appt.Location = [char]44536 + [char]47004 + [char]46300 + " " + [char]51064 + [char]53552 + [char]53080 + [char]54000 + [char]45340 + [char]53448 + [char]53448 + " " + [char]49436 + [char]50872 + " " + [char]54028 + [char]47476 + [char]45208 + [char]49828
    $appt.Body = "SAP Build를 활용한 비즈니스 프로세스 자동화 및 앱 개발 전략 세미나 초청 건입니다."
    
    # 시간 설정: 2026-03-19 14:00 ~ 17:30
    $appt.Start = Get-Date -Year 2026 -Month 3 -Day 19 -Hour 14 -Minute 0 -Second 0
    $appt.End = Get-Date -Year 2026 -Month 3 -Day 19 -Hour 17 -Minute 30 -Second 0
    
    # 미리 알림 설정: 1일 전(1440분)
    # 구글 캘린더 동기화 최적화를 위해 첫 번째 알림을 Outlook 기본값으로 설정
    $appt.ReminderSet = $true
    $appt.ReminderMinutesBeforeStart = 1440 
    
    $appt.Save()
    
    # 3시간 전 알림은 구글 캘린더 동기화 후 수동 확인이 필요할 수 있으나 
    # 일단 Outlook에 1일 전(최장 시간) 알림으로 등록하여 안전하게 동기화 유도
    
    "SUCCESS"
} catch {
    $_.Exception.Message
}
