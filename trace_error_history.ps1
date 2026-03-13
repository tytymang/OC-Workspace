
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $targetSubject = "Error ZSCMR3000"
    $found = $false
    
    # 최근 100개 중 동일 주제 찾기
    for ($i = 1; $i -le 100; $i++) {
        $item = $items.Item($i)
        if ($item.Subject -like "*$targetSubject*") {
            Write-Output "Sent: $($item.ReceivedTime) | From: $($item.SenderName)"
        }
    }
} catch {
    Write-Output $_.Exception.Message
}
