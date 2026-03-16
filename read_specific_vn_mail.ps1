
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    foreach ($m in $items) {
        if ($m.Subject -match "VN" -and $m.ReceivedTime.ToString("HH:mm") -eq "10:37") {
            Write-Host "SUBJECT: $($m.Subject)"
            Write-Host "BODY: $($m.Body)"
            break
        }
    }
} catch {
    Write-Error $_.Exception.Message
}
