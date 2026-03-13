$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    # 리스트에서 첫 번째 항목이 이수정 님 메일이었으므로 바로 추출
    $item = $items.Item(1)
    
    $results = [PSCustomObject]@{
        Sender = $item.SenderName
        Subject = $item.Subject
        Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        Body = $item.Body
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}