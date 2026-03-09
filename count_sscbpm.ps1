$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6)

$unread = $Inbox.Items | Where-Object { $_.Unread -eq $true }
$sscbpm_unread = @($unread | Where-Object { $_.SenderName -match "SSCBPM" })

$all_sscbpm = @($Inbox.Items | Where-Object { $_.SenderName -match "SSCBPM" })

[PSCustomObject]@{
    TotalSSCBPM = $all_sscbpm.Count
    UnreadSSCBPM = $sscbpm_unread.Count
} | ConvertTo-Json
