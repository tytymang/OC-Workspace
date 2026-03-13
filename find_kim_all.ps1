$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 전체 메일 중 보낸 사람에 "김하영"이 포함된 것 검색
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    for ($i = 1; $i -le 200; $i++) {
        $item = $items.Item($i)
        if ($item.SenderName -like "*김하영*") {
            $results += [PSCustomObject]@{
                Index = $i
                Sender = $item.SenderName
                Subject = $item.Subject
                Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            }
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}