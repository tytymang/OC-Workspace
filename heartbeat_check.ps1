$ErrorActionPreference = "Continue"
try {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    $now = Get-Date
    $twoHoursLater = $now.AddHours(2)
    
    # 1. Check Calendar (Next 2 hours)
    $calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar
    $items = $calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
    
    $filter = "[Start] >= '$($now.ToString("yyyy-MM-dd HH:mm"))' AND [Start] <= '$($twoHoursLater.ToString("yyyy-MM-dd HH:mm"))'"
    $upcomingMeetings = $items.Restrict($filter)
    
    $meetingReports = @()
    foreach ($m in $upcomingMeetings) {
        $meetingReports += "일정: $($m.Subject) ($($m.Start.ToString("HH:mm")))"
    }

    # 2. Check for NEW VIP emails (since 07:44)
    # VIPs: 이정훈, 이정우, 이상무, 이영주
    $inbox = $namespace.GetDefaultFolder(6)
    $vipNames = @("이정훈", "이정우", "이상무", "이영주")
    $newVipEmails = @()
    
    foreach ($name in $vipNames) {
        $filter = "[UnRead] = true AND [ReceivedTime] > '2026-03-25 07:44' AND [SenderName] = '$name'"
        $found = $inbox.Items.Restrict($filter)
        foreach ($e in $found) {
            $newVipEmails += "[$($e.SenderName)] $($e.Subject)"
        }
    }

    $res = @{
        meetings = $meetingReports
        vipEmails = $newVipEmails
    }
    $res | ConvertTo-Json
} catch {
    "ERROR: $($_.Exception.Message)"
}
