$ErrorActionPreference = "Stop"
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$results = @{
    emails = @()
    calendar = @()
}

$now = Get-Date
$twoHoursLater = $now.AddHours(2)

# VIP Names using Unicode char points to avoid encoding issues
$vip1 = -join @([char]0xC774, [char]0xC815, [char]0xD6C8) # 이정훈
$vip2 = -join @([char]0xC774, [char]0xC815, [char]0xC6B0) # 이정우
$vip3 = -join @([char]0xC774, [char]0xC0C1, [char]0xBB34) # 이상무
$vip4 = -join @([char]0xC774, [char]0xC601, [char]0xC91C) # 이영주
$vipNames = @($vip1, $vip2, $vip3, $vip4)

function Scan-VIP-Emails($folder) {
    try {
        $items = $folder.Items.Restrict("[UnRead] = true")
        foreach ($item in $items) {
            $sender = try { $item.SenderName } catch { "" }
            foreach ($vip in $vipNames) {
                if ($sender -like "*$vip*") {
                    $results.emails += @{
                        Sender = $sender
                        Subject = $item.Subject
                        Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                    }
                    break
                }
            }
        }
    } catch {}
    
    # Simple recursion for Inbox subfolders
    if ($folder.Name -match "Inbox" -or $folder.Name -match "받은") {
        foreach ($sub in $folder.Folders) { try { Scan-VIP-Emails $sub } catch {} }
    }
}

foreach ($root in $namespace.Folders) {
    try {
        $inbox = $root.Folders.Item("Inbox")
        if ($null -ne $inbox) { Scan-VIP-Emails $inbox }
    } catch {
        try {
            # Try Korean Inbox name
            $krInbox = $root.Folders | Where-Object { $_.Name -match -join @([char]0xBC1B, [char]0xC740, [char]0x0020, [char]0xD3B8, [char]0xC9C0, [char]0xD568) }
            if ($null -ne $krInbox) { Scan-VIP-Emails $krInbox }
        } catch {}
    }
}

# 2. Calendar Check (Next 2 hours)
try {
    $calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar
    $items = $calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
    # For restrict filter, using string format
    $startTimeStr = $now.ToString("yyyy-MM-dd HH:mm")
    $endTimeStr = $twoHoursLater.ToString("yyyy-MM-dd HH:mm")
    $filter = "[Start] >= '$startTimeStr' AND [Start] <= '$endTimeStr'"
    $upcoming = $items.Restrict($filter)

    foreach ($item in $upcoming) {
        $results.calendar += @{
            Subject = $item.Subject
            Start = $item.Start.ToString("yyyy-MM-dd HH:mm")
        }
    }
} catch {}

$results | ConvertTo-Json -Compress
