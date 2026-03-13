
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    for ($i = 1; $i -le 30; $i++) {
        $item = $items.Item($i)
        if ($item.UnRead) {
            $results += [PSCustomObject]@{
                ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                SenderName   = $item.SenderName
                Subject      = $item.Subject
            }
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output $_.Exception.Message
}
