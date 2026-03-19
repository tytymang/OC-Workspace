try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $unreadItems = $inbox.Items.Restrict("[UnRead] = true")

    $results = @()
    foreach ($item in $unreadItems) {
        $results += @{
            ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Sender = $item.SenderName
            Subject = $item.Subject
        }
    }
    $results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\all_unread_emails.json" -Encoding UTF8
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\ps_error.log"
}
