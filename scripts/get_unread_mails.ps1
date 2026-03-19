$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)
$unreadItems = $inbox.Items.Restrict("[UnRead] = true")
$unreadItems.Sort("[ReceivedTime]", $true)

$results = @()
$count = 0
foreach ($item in $unreadItems) {
    if ($count -ge 10) { break }
    
    $summaryText = "System-generated notification"
    if ($item.Subject -match "RE:") {
        $summaryText = "Reply to previous conversation"
    }

    $replyRequested = "No"
    # Hex check for common request keywords if possible, but simpler check for now:
    if ($item.Subject -like "*회신*" -or $item.Subject -like "*답변*" -or $item.Subject -like "*제출*") {
        $replyRequested = "Yes"
    }

    $results += [PSCustomObject]@{
        Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        Sender   = $item.SenderName
        Subject  = $item.Subject
        Summary  = $summaryText
        ReplyReq = $replyRequested
    }
    $count++
}

$results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\unread_mails.json" -Encoding Unicode
