
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar
    
    # 1. 출국 일정 (3/18 10:05 - 13:25)
    $dep = $outlook.CreateItem(1) # olAppointmentItem
    $dep.Subject = "[출국] 최현구 상무님 VN 출장 (VN0417)"
    $dep.Location = "ICN -> HAN"
    $dep.Start = "2026-03-18 10:05"
    $dep.End = "2026-03-18 13:25"
    $dep.Body = "Trip.com 예약 확정 정보 기반`n- 편명: VN0417`n- 14:30 픽업 대기 (Mr Chuyen, 29H 934.95)"
    $dep.ReminderSet = $true
    $dep.ReminderMinutesBeforeStart = 1440 # 1 day
    $dep.Save()
    
    # 2. 귀국 일정 (4/02 23:35 - 4/03 05:50)
    $ret = $outlook.CreateItem(1)
    $ret.Subject = "[귀국] 최현구 상무님 VN 출장 (VN0416)"
    $ret.Location = "HAN -> ICN"
    $ret.Start = "2026-04-02 23:35"
    $ret.End = "2026-04-03 05:50"
    $ret.Body = "Trip.com 예약 확정 정보 기반`n- 편명: VN0416`n- 19:10 숙소에서 공항 이동 시작"
    $ret.ReminderSet = $true
    $ret.ReminderMinutesBeforeStart = 1440
    $ret.Save()

    Write-Host "SUCCESS: 2 appointments registered with clean Korean encoding."
} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($null -ne $outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}
