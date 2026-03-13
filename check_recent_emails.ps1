
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6) # olFolderInbox
$targetDate = (Get-Date).AddDays(-1).Date

$emails = $inbox.Items | Where-Object { $_.ReceivedTime -ge $targetDate } | Sort-Object ReceivedTime -Descending

$results = @()
foreach ($mail in $emails) {
    $results += [PSCustomObject]@{
        ReceivedTime = $mail.ReceivedTime
        SenderName   = $mail.SenderName
        Subject      = $mail.Subject
        UnRead       = $mail.UnRead
    }
}

$results | ConvertTo-Json | Out-File -FilePath "recent_emails.json" -Encoding UTF8
