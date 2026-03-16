
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    
    foreach ($m in $items) {
        if ($m.SenderName -like "*Choi, ByungSoon*" -and $m.ReceivedTime -gt (Get-Date 2026-03-10)) {
            Write-Output "--- MAIL START ---"
            Write-Output "SUBJECT: $($m.Subject)"
            Write-Output "BODY: $($m.Body.Substring(0, [Math]::Min(1000, $m.Body.Length)))"
            Write-Output "--- MAIL END ---"
        }
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
