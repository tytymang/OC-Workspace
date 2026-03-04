$OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace('MAPI')
$Inbox = $Namespace.GetDefaultFolder(6)
$UnreadItems = $Inbox.Items.Restrict("[Unread] = true")

if ($UnreadItems.Count -eq 0) {
    Write-Output "읽지 않은 메일이 없습니다."
} else {
    $Results = @()
    foreach ($Item in $UnreadItems) {
        $Obj = [PSCustomObject]@{
            Subject = $Item.Subject
            Sender = $Item.SenderName
            ReceivedTime = $Item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Body = if ($Item.Body.Length -gt 500) { $Item.Body.Substring(0, 500) } else { $Item.Body }
        }
        $Results += $Obj
    }
    $Results | ConvertTo-Json
}
