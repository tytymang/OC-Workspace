
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6) # olFolderInbox
    $deletedItems = $namespace.GetDefaultFolder(3) # olFolderDeletedItems
    
    $today = Get-Date -Hour 0 -Minute 0 -Second 0
    $filter = "[ReceivedTime] >= '$($today.ToString("g"))'"
    $items = $inbox.Items.Restrict($filter)
    
    $count = $items.Count
    $successCount = 0
    
    # We must loop backwards when moving/deleting items
    for ($i = $count; $i -ge 1; $i--) {
        $item = $items.Item($i)
        $item.Move($deletedItems) | Out-Null
        $successCount++
    }
    
    Write-Output "SUCCESS: Moved $successCount emails to Deleted Items."
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
