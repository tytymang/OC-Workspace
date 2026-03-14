$ErrorActionPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# 1. Check Emails (Last 30 mins)
$Inbox = $Namespace.GetDefaultFolder(6) 
$ThirtyMinsAgo = (Get-Date).AddMinutes(-30)
$Filter = "[ReceivedTime] >= '" + $ThirtyMinsAgo.ToString("g") + "'"
$RecentEmails = $Inbox.Items.Restrict($Filter)

$VipEmails = @()
foreach ($Mail in $RecentEmails) {
    if ($Mail.Importance -eq 2 -or $Mail.SenderName -match "이정우|나여나") {
        $VipEmails += $Mail.SenderName + ": " + $Mail.Subject
    }
}

# 2. Check Calendar (Next 2 hours)
$Calendar = $Namespace.GetDefaultFolder(9) 
$TwoHoursLater = (Get-Date).AddHours(2)
$NowStr = (Get-Date).ToString("g")
$LaterStr = $TwoHoursLater.ToString("g")
$CalFilter = "[Start] >= '" + $NowStr + "' AND [Start] <= '" + $LaterStr + "'"
$UpcomingEvents = $Calendar.Items.Restrict($CalFilter)
$Events = @()
foreach ($Event in $UpcomingEvents) {
    $Events += $Event.Start.ToString("HH:mm") + " - " + $Event.Subject
}

# 3. Check Git Sync
$GitStatus = git status --porcelain
$SyncNeeded = [bool]$GitStatus

# Output Results
Write-Output "---RESULTS---"
Write-Output "VIP_EMAILS:"
$VipEmails | ForEach-Object { Write-Output $_ }
Write-Output "EVENTS:"
$Events | ForEach-Object { Write-Output $_ }
Write-Output "SYNC_NEEDED: $SyncNeeded"
