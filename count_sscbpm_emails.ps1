$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6) # 6 is olFolderInbox

# Restrict to emails where SenderName contains SSCBPM or exactly matches
$Filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0042001E"" like '%SSCBPM%' OR ""http://schemas.microsoft.com/mapi/proptag/0x0C1A001E"" like '%SSCBPM%'"
# Let's just use a simpler filter or iterate if restrict fails, but iteration on huge inbox is slow.
# Simple restrict:
$Filter2 = "[SenderName] = 'SSCBPM'"

$SSCBPM_Emails = $Inbox.Items.Restrict($Filter2)
$totalCount = $SSCBPM_Emails.Count

$unreadCount = 0
foreach ($email in $SSCBPM_Emails) {
    if ($email.Unread) {
        $unreadCount++
    }
}

[PSCustomObject]@{
    TotalCount = $totalCount
    UnreadCount = $unreadCount
} | ConvertTo-Json
