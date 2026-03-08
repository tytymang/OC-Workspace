$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6) # 6 is olFolderInbox

$UnreadEmails = $Inbox.Items | Where-Object { $_.Unread -eq $true } | Sort-Object ReceivedTime -Descending | Select-Object -First 10

$results = @()
foreach ($email in $UnreadEmails) {
    $body = $email.Body
    if ($body -ne $null) {
        $body = $body.Trim() -replace '\s+', ' '
        $body = $body.Substring(0, [math]::Min($body.Length, 300))
    } else {
        $body = ""
    }
    
    $results += [PSCustomObject]@{
        ReceivedTime = $email.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        SenderName = $email.SenderName
        Subject = $email.Subject
        Body = $body
    }
}

$results | ConvertTo-Json -Depth 3 -Compress
