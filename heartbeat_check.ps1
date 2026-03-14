$ErrorActionPreference = "SilentlyContinue"
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# 1. Check Emails (Last 30 mins)
$Inbox = $Namespace.GetDefaultFolder(6) # olFolderInbox
$ThirtyMinsAgo = (Get-Date).AddMinutes(-30)
$Filter = "[ReceivedTime] >= '$($ThirtyMinsAgo.ToString("g"))'"
$RecentEmails = $Inbox.Items.Restrict($Filter)

$VipEmails = @()
foreach ($Mail in $RecentEmails) {
    if ($Mail.Importance -eq 2 -or $Mail.SenderName -match "이정우|나여나") {
        $VipEmails += "$($Mail.SenderName): $($Mail.Subject)"
    }
}

# 2. Check Calendar (Next 2 hours)
$Calendar = $Namespace.GetDefaultFolder(9) # olFolderCalendar
$TwoHoursLater = (Get-Date).AddHours(2)
$CalFilter = "[Start] >= '$((Get-Date).ToString("g"))' AND [Start] <= '$($TwoHoursLater.ToString("g"))'"
$UpcomingEvents = $Calendar.Items.Restrict($CalFilter)
$Events = @()
foreach ($Event in $UpcomingEvents) {
    $Events += "$($Event.Start.ToString("HH:mm")) - $($Event.Subject)"
}

# 3. Check Git Sync
$GitStatus = git status --porcelain
$SyncNeeded = $false
if ($GitStatus) { $SyncNeeded = $true }

# Output Results
$Result = @{
    VipEmails = $VipEmails
    Events = $Events
    SyncNeeded = $SyncNeeded
}

$Result | ConvertTo-Json
