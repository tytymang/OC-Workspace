$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6)

$UnreadEmails = $Inbox.Items | Where-Object { $_.Unread -eq $true } | Sort-Object ReceivedTime -Descending

$out = @()
foreach ($email in $UnreadEmails) {
    if ($email.Importance -eq 2) { # High Importance
        $out += [PSCustomObject]@{
            ReceivedTime = $email.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            SenderName = $email.SenderName
            Subject = $email.Subject
            Importance = "High"
        }
    } else {
        $out += [PSCustomObject]@{
            ReceivedTime = $email.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            SenderName = $email.SenderName
            Subject = $email.Subject
            Importance = "Normal"
        }
    }
}
$out | Select-Object -First 5 | ConvertTo-Json
