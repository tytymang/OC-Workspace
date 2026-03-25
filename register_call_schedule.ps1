$ErrorActionPreference = "Stop"
try {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $item = $outlook.CreateItem(1) # olAppointmentItem

    # Title: "데이터 솔루션 정우찬 주임 통화"
    # [char]0xB370 (데), 0xCC21 (이), 0xD130 (터), 0x0020 ( ), 0xC124 (솔), 0xB8E8 (션), 0x0020 ( ), 0xC815 (정), 0xC6B0 (우), 0xCC2C (찬), 0x0020 ( ), 0xC91C (주), 0xC784 (임), 0x0020 ( ), 0xD1B5 (통), 0xD654 (화)
    $subjectParts = @(
        [char]0xB370, [char]0xCC21, [char]0xD130, [char]0x0020, 
        [char]0xC124, [char]0xB8E8, [char]0x0020, [char]0xC124, [char]0xB8E8, [char]0xC540, # 솔루션 오타 교정 시도 대신 정확히 조립
        [char]0xC124, [char]0xB8E8, [char]0xC520 # 솔루션 (솔:0xC124, 루:0xB8E8, 션:0xC120)
    )
    # 정정: 데이터(B370 CC21 D130) 솔루션(C124 B8E8 C120) 정우찬(C815 C6B0 CC2C) 주임(C91C C784) 통화(D1B5 D654)
    $subject = -join @(
        [char]0xB370, [char]0xCC21, [char]0xD130, [char]0x0020, # 데이터 
        [char]0xC124, [char]0xB8E8, [char]0xC120, [char]0x0020, # 솔루션 
        [char]0xC815, [char]0xC6B0, [char]0xCC2C, [char]0x0020, # 정우찬 
        [char]0xC91C, [char]0xC784, [char]0x0020,             # 주임 
        [char]0xD1B5, [char]0xD654                            # 통화
    )

    $startTime = Get-Date -Year 2026 -Month 3 -Day 25 -Hour 11 -Minute 0 -Second 0
    $endTime = $startTime.AddMinutes(30)

    $item.Subject = $subject
    $item.Start = $startTime
    $item.End = $endTime
    $item.ReminderSet = $true
    $item.ReminderMinutesBeforeStart = 15 # Default 15 min reminder
    
    # Force Google Sync Reminder (17 hours before? No, this is a normal meeting today)
    # But skill-office-automation says Google Sync Strategy:
    # If it was for tomorrow, I would set 1020 mins. For today, 15 mins is fine.
    
    $item.Display()
    
    Write-Host "SUCCESS: Appointment window displayed. Subject: $subject"
} catch {
    Write-Error "FAILED: $($_.Exception.Message)"
}
