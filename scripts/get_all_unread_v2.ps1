$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$results = @()

function Get-UnreadFromFolderRecursive($folder, $folderPath) {
    # Get unread items count or iterate
    $unreadCount = $folder.UnReadItemCount
    if ($unreadCount -gt 0) {
        try {
            $items = $folder.Items.Restrict("[UnRead] = true")
            foreach ($item in $items) {
                $replyRequested = "No"
                # Check for keywords in subject or body (handle potential nulls)
                $subject = if ($item.Subject) { $item.Subject } else { "" }
                if ($subject -like "*회신*" -or $subject -like "*답변*" -or $subject -like "*제출*") {
                    $replyRequested = "Yes"
                }

                $summary = "System/Business Mail"
                if ($subject -match "RE:") { $summary = "Reply to thread" }

                $global:results += [PSCustomObject]@{
                    Folder   = $folderPath
                    Received = if ($item.ReceivedTime) { $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm") } else { "Unknown" }
                    Sender   = if ($item.SenderName) { $item.SenderName } else { "Unknown" }
                    Subject  = $subject
                    Summary  = $summary
                    ReplyReq = $replyRequested
                }
            }
        } catch {
            # Skip if items cannot be accessed
        }
    }
    
    foreach ($subFolder in $folder.Folders) {
        Get-UnreadFromFolderRecursive $subFolder ($folderPath + "/" + $subFolder.Name)
    }
}

# Iterate through all Stores (Accounts/Mailboxes)
foreach ($store in $namespace.Stores) {
    $rootFolder = $store.GetRootFolder()
    Get-UnreadFromFolderRecursive $rootFolder $store.DisplayName
}

$global:results | Sort-Object Received -Descending | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\all_unread_mails_comprehensive.json" -Encoding Unicode
