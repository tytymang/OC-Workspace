
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # Filter: Unread emails received today (2026-03-05)
    $todayStr = Get-Date -Year 2026 -Month 3 -Day 5 -Hour 0 -Minute 0 -Second 0 -Format "yyyy-MM-dd HH:mm"
    $filter = "[UnRead] = true AND [ReceivedTime] >= '$todayStr'"
    $items = $inbox.Items.Restrict($filter)
    $items.Sort("[ReceivedTime]", $true)
    
    $maxToDelete = [Math]::Min($items.Count, 5)
    $deletedSubjects = @()
    
    # To delete items from a collection, it's safer to use a fixed list of entry IDs or similar
    # or just loop through a fixed number of times always taking the first one if sorting is maintained.
    # However, for simplicity and safety, we will collect the items first.
    $toDelete = @()
    for ($i = 1; $i -le $maxToDelete; $i++) {
        $toDelete += $items.Item($i)
    }

    foreach ($mail in $toDelete) {
        $sub = $mail.Subject
        $mail.Delete()
        $deletedSubjects += $sub
    }
    
    $deletedSubjects | ConvertTo-Json
} catch {
    $_.Exception.Message | Out-File -FilePath "error.log"
    throw $_
}
