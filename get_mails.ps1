
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    for ($i=1; $i -le 5; $i++) {
        $m = $items.Item($i)
        Write-Output "MAIL: [$($m.ReceivedTime.ToString('MM/dd HH:mm'))] $($m.SenderName): $($m.Subject)"
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
