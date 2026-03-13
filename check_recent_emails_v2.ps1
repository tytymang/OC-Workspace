
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $targetDate = (Get-Date).AddDays(-1)

    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)

    $results = @()
    $count = 0
    foreach ($item in $items) {
        if ($item.ReceivedTime -lt $targetDate -or $count -ge 10) { break }
        $results += [PSCustomObject]@{
            ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
            SenderName   = $item.SenderName
            Subject      = $item.Subject
            UnRead       = $item.UnRead
        }
        $count++
    }

    $results | ConvertTo-Json | Out-File -FilePath "recent_emails.json" -Encoding utf8
} catch {
    $_.Exception.Message | Out-File -FilePath "error_log.txt"
}
