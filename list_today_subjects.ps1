
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $today = (Get-Date).Date
    foreach ($m in $items) {
        if ($m.ReceivedTime.Date -lt $today) { break }
        Write-Host "[$($m.ReceivedTime.ToString('HH:mm'))] Subject: $($m.Subject)"
    }
} catch {
    Write-Error $_.Exception.Message
}
