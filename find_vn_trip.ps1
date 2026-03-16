
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    
    foreach ($m in $items) {
        # Search for "VN" and "출장" (or part of it)
        if ($m.Subject -match "VN" -and $m.Subject -match "출장") {
            Write-Host "---"
            Write-Host "Subject: $($m.Subject)"
            Write-Host "Received: $($m.ReceivedTime)"
            Write-Host "Body: $($m.Body)"
        }
    }
} catch {
    Write-Error $_.Exception.Message
}
