
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    
    foreach ($m in $items) {
        if ($m.Subject -eq "RE: VN  û") {
            Write-Output "SUBJECT: $($m.Subject)"
            Write-Output "BODY: $($m.Body)"
            break
        }
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
