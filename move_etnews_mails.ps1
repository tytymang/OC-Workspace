
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
    $inbox = $namespace.GetDefaultFolder(6) # olFolderInbox
    $deletedItems = $namespace.GetDefaultFolder(3) # olFolderDeletedItems
    
    # "전자신문" 키워드로 검색 (보낸 사람 이름 또는 제목)
    $items = $inbox.Items
    $count = 0
    
    # 컬렉션을 직접 반복하면서 삭제하면 인덱스가 꼬이므로 리스트에 먼저 담음
    $targetItems = @()
    foreach ($item in $items) {
        if ($item.SenderName -like "*전자신문*" -or $item.Subject -like "*전자신문*") {
            $targetItems += $item
        }
    }
    
    foreach ($item in $targetItems) {
        $subject = $item.Subject
        $item.Move($deletedItems) | Out-Null
        Write-Output "Moved to Deleted Items: $subject"
        $count++
    }
    
    Write-Output "Total $count messages moved."
} catch {
    Write-Error $_.Exception.Message
}
