
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $deleted = $namespace.GetDefaultFolder(3)
    
    # 받은 편지함에서 SSCBPM 메일 찾아 삭제(지운 편지함 이동)
    $items = $inbox.Items
    $count = $items.Count
    $deletedCount = 0
    
    for ($i = $count; $i -ge 1; $i--) {
        $item = $items.Item($i)
        if ($item.SenderName -like "*SSCBPM*" -or $item.Subject -like "*SSCBPM*") {
            $item.Move($deleted) | Out-Null
            $deletedCount++
        }
    }
    
    Write-Output "SUCCESS: Moved $deletedCount SSCBPM emails from Inbox to Deleted Items."
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
