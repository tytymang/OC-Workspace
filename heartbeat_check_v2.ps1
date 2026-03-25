$ErrorActionPreference = "Continue"
try {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    $now = Get-Date
    $twoHoursLater = $now.AddHours(2)
    
    # Check Calendar
    $calendar = $namespace.GetDefaultFolder(9)
    $items = $calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
    $filter = "[Start] >= '$($now.ToString("yyyy-MM-dd HH:mm"))' AND [Start] <= '$($twoHoursLater.ToString("yyyy-MM-dd HH:mm"))'"
    $upcoming = $items.Restrict($filter)
    
    $meetings = @()
    foreach ($m in $upcoming) {
        $meetings += "$($m.Subject) ($($m.Start.ToString("HH:mm")))"
    }

    # Check for NEW unread emails since 07:44
    $inbox = $namespace.GetDefaultFolder(6)
    $unreadItems = $inbox.Items.Restrict("[UnRead] = true")
    $newEmails = @()
    foreach ($e in $unreadItems) {
        if ($e.ReceivedTime -gt [datetime]"2026-03-25 07:44:00") {
            $newEmails += "$($e.SenderName): $($e.Subject)"
        }
    }

    @{ meetings = $meetings; newEmails = $newEmails } | ConvertTo-Json
} catch {
    "ERROR: $($_.Exception.Message)"
}
