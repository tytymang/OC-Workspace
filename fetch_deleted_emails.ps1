
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $deletedItemsFolder = $namespace.GetDefaultFolder(3) # 3 is OlDefaultFolders.olFolderDeletedItems
    $items = $deletedItemsFolder.Items
    
    # Sort by received time descending
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    # Get top 10 deleted emails
    $count = [Math]::Min($items.Count, 10)
    for ($i = 1; $i -le $count; $i++) {
        $item = $items.Item($i)
        
        # Check if it's a MailItem (some items in Deleted Items might be different types)
        if ($item.MessageClass -eq "IPM.Note") {
            $results += [PSCustomObject]@{
                Time = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                Sender = $item.SenderName
                Subject = $item.Subject
            }
        }
    }
    
    $results | ConvertTo-Json
} catch {
    Write-Error $_.Exception.Message
}
