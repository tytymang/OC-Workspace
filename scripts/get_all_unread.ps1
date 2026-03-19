$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

$results = @()

function Get-UnreadFromFolder($folder, $folderPath) {
    $items = $folder.Items.Restrict("[UnRead] = true")
    foreach ($item in $items) {
        $replyRequested = "No"
        if ($item.Subject -like "*회신*" -or $item.Subject -like "*답변*" -or $item.Subject -like "*제출*") {
            $replyRequested = "Yes"
        }

        $summary = "System/Business Mail"
        if ($item.Subject -match "RE:") { $summary = "Reply to thread" }

        $global:results += [PSCustomObject]@{
            Folder   = $folderPath
            Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Sender   = $item.SenderName
            Subject  = $item.Subject
            Summary  = $summary
            ReplyReq = $replyRequested
        }
    }
    
    foreach ($subFolder in $folder.Folders) {
        Get-UnreadFromFolder $subFolder ($folderPath + "/" + $subFolder.Name)
    }
}

Get-UnreadFromFolder $inbox "Inbox"

$global:results | Sort-Object Received -Descending | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\all_unread_mails.json" -Encoding Unicode
