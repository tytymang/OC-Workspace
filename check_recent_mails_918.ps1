$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)
$unreadItems = $inbox.Items.Restrict("[UnRead] = true")
$unreadItems.Sort("[ReceivedTime]", $true)

$results = @()
foreach ($item in $unreadItems) {
    if ($item.ReceivedTime -gt (Get-Date).AddMinutes(-30)) {
        $results += @{
            ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Sender = $item.SenderName
            Subject = $item.Subject
        }
    }
}
$results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\heartbeat_recent_mails_918.json" -Encoding UTF8
