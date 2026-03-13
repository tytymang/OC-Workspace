$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # "비서/기획" 폴더 접근
    $secretaryFolder = $inbox.Folders.Item(7)
    # "김하영 수석" 폴더 접근
    $kimFolder = $secretaryFolder.Folders.Item(2)
    
    # 폴더 내 모든 메일 제목 출력하여 수동 확인
    $items = $kimFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    $max = if ($items.Count -lt 50) { $items.Count } else { 50 }
    for ($i = 1; $i -le $max; $i++) {
        $item = $items.Item($i)
        $results += [PSCustomObject]@{
            Index = $i
            Sender = $item.SenderName
            Subject = $item.Subject
            Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            HasAttachments = ($item.Attachments.Count -gt 0)
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}