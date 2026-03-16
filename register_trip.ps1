
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
    
    # 1. 출국 일정 등록 (3/18 10:00)
    $departure = $outlook.CreateItem(1) # olAppointmentItem
    $departure.Subject = "[출국] VN 출장 (ICN -> HAN)"
    $departure.Start = "2026-03-18 10:00"
    $departure.Duration = 300 # 5시간 비행 가정
    $departure.ReminderSet = $true
    # 알림 설정: 1일 전 (1440분)
    $departure.ReminderMinutesBeforeStart = 1440
    $departure.Save()
    
    # 2. 귀국 일정 등록 (4/2 23:00)
    $arrival = $outlook.CreateItem(1)
    $arrival.Subject = "[귀국] VN 출장 (HAN -> ICN)"
    $arrival.Start = "2026-04-02 23:00"
    $arrival.Duration = 300
    $arrival.ReminderSet = $true
    $arrival.ReminderMinutesBeforeStart = 1440
    $arrival.Save()

    Write-Output "SUCCESS: Registered 2 trip events to Outlook Calendar."
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
