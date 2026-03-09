$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6) # olFolderInbox

$UnreadEmails = $Inbox.Items | Where-Object { $_.Unread -eq $true }

$deletedCount = 0
foreach ($email in $UnreadEmails) {
    if ($email.Subject -match "\[Today's Birthday\]") {
        $subject = $email.Subject
        $email.Delete()
        Write-Output "Deleted: $subject"
        $deletedCount++
    }
}

if ($deletedCount -eq 0) {
    Write-Output "No birthday emails found to delete."
} else {
    Write-Output "Total deleted: $deletedCount"
}
