
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6) # 6 is OlDefaultFolders.olFolderInbox
    $unreadItems = $inbox.Items.Restrict("[UnRead] = true")
    
    # Sort by received time descending
    $unreadItems.Sort("[ReceivedTime]", $true)
    
    $results = @()
    # Get top 10 unread emails
    $count = [Math]::Min($unreadItems.Count, 10)
    for ($i = 1; $i -le $count; $i++) {
        $item = $unreadItems.Item($i)
        $body = $item.Body
        if ($body -eq $null) { $body = "" }
        $summary = $body.Trim()
        if ($summary.Length -gt 100) {
            $summary = $summary.Substring(0, 100) + "..."
        }
        
        $results += [PSCustomObject]@{
            Time = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Sender = $item.SenderName
            Subject = $item.Subject
            Summary = $summary
        }
    }
    
    $results | ConvertTo-Json
} catch {
    Write-Error $_.Exception.Message
}
