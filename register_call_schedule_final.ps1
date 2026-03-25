$ErrorActionPreference = "Stop"
try {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $item = $outlook.CreateItem(1) # olAppointmentItem

    # Title: "데이터 솔루션 정우찬 주임 통화"
    # [char]0xB370 (데), 0xCC21 (이), 0xD130 (터), 0x0020 ( ), 0xC124 (솔), 0xB8E8 (션) 
    # 정:0xC815, 우:0xC6B0, 찬:0xCC2C, 주:0xC91C, 임:0xC784, 통:0xD1B5, 화:0xD654
    $subject = -join @(
        [char]0xB370, [char]0xCC21, [char]0xD130, [char]0x0020, 
        [char]0xC124, [char]0xB8E8, [char]0xC120, [char]0x0020, 
        [char]0xC815, [char]0xC6B0, [char]0xCC2C, [char]0x0020, 
        [char]0xC91C, [char]0xC784, [char]0x0020, 
        [char]0xD1B5, [char]0xD654
    )

    $startTime = Get-Date -Year 2026 -Month 3 -Day 25 -Hour 11 -Minute 0 -Second 0
    $endTime = $startTime.AddMinutes(30)

    $item.Subject = $subject
    $item.Start = $startTime
    $item.Duration = 30
    $item.ReminderSet = $true
    $item.ReminderMinutesBeforeStart = 15
    
    $item.Display()
    
    # Force Sync (Skill-Strategy)
    # $namespace.SyncObjects.Item(1).Start()
    
} catch {
    Write-Error "FAILED: $($_.Exception.Message)"
}
