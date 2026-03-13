
try {
    Write-Output "Starting Outlook script"
    $outlook = New-Object -ComObject Outlook.Application
    Write-Output "Outlook object created"
    $ns = $outlook.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    Write-Output "Inbox folder: $($inbox.Name)"
    
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    for ($i = 1; $i -le 5; $i++) {
        $item = $items.Item($i)
        $results += [PSCustomObject]@{
            ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
            SenderName   = $item.SenderName
            Subject      = $item.Subject
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Error $_.Exception.Message
}
