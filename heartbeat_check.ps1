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

# 1. VIP Email Check (Top-level Inbox only for quick scan or recursive if small)
$vipNames = @("이정훈", "이정우", "이상무", "이영주")

function Scan-VIP-Emails($folder) {
    try {
        $items = $folder.Items.Restrict("[UnRead] = true")
        foreach ($item in $items) {
            foreach ($vip in $vipNames) {
                if ($item.SenderName -like "*$vip*") {
                    $results.emails += @{
                        Sender = $item.SenderName
                        Subject = $item.Subject
                        Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                    }
                    break
                }
            }
        }
    } catch {}
    # Scan subfolders if they look like Inbox subfolders
    if ($folder.Name -match "Inbox" -or $folder.Name -match "받은") {
        foreach ($sub in $folder.Folders) { Scan-VIP-Emails $sub }
    }
}

foreach ($root in $namespace.Folders) {
    Scan-VIP-Emails $root.Folders.Item("Inbox")
}

# 2. Calendar Check (Next 2 hours)
$calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar
$items = $calendar.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$filter = "[Start] >= '$($now.ToString("g"))' AND [Start] <= '$($twoHoursLater.ToString("g"))'"
$upcoming = $items.Restrict($filter)

foreach ($item in $upcoming) {
    $results.calendar += @{
        Subject = $item.Subject
        Start = $item.Start.ToString("yyyy-MM-dd HH:mm")
    }
}

$results | ConvertTo-Json -Compress
