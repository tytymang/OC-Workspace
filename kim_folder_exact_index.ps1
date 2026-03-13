$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $secretaryFolder = $inbox.Folders.Item(7)
    $kimFolder = $secretaryFolder.Folders.Item(2)
    
    $items = $kimFolder.Items
    $results = @()
    # 상위 5개 메일의 인덱스와 제목을 정확히 대조
    for ($i = 1; $i -le 5; $i++) {
        $item = $items.Item($i)
        $results += [PSCustomObject]@{
            RealIndex = $i
            Subject = $item.Subject
            Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            HasAttachments = ($item.Attachments.Count -gt 0)
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}