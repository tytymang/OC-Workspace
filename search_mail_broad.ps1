
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6) # Inbox check in case it's an internal reply
    $sent = $namespace.GetDefaultFolder(5) # Sent
    
    $found = $null
    
    # Check Sent first (Last 100)
    $countSent = [Math]::Min($sent.Items.Count, 100)
    for ($i = $sent.Items.Count; $i -gt ($sent.Items.Count - $countSent); $i--) {
        $item = $sent.Items.Item($i)
        if ($item.Subject -like "*2월*마감*완료*확인*") {
            $found = $item
            break
        }
    }
    
    if ($found -eq $null) {
        # Check Inbox (Last 100)
        $countInbox = [Math]::Min($inbox.Items.Count, 100)
        for ($i = $inbox.Items.Count; $i -gt ($inbox.Items.Count - $countInbox); $i--) {
            $item = $inbox.Items.Item($i)
            if ($item.Subject -like "*2월*마감*완료*확인*") {
                $found = $item
                break
            }
        }
    }
    
    if ($found -ne $null) {
        $recipients = @()
        foreach ($recipient in $found.Recipients) {
            # Type 1 = olTo, Type 2 = olCC, Type 3 = olBCC
            if ($recipient.Type -eq 1) {
                $recipients += $recipient.Name
            }
        }
        $recipients | ConvertTo-Json
    } else {
        "NOT_FOUND"
    }
} catch {
    $_.Exception.Message | Out-File -FilePath "error_search.log"
    throw $_
}
