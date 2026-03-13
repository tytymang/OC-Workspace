$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    # 500개까지 뒤져보자
    $limit = if ($items.Count -lt 500) { $items.Count } else { 500 }
    
    for ($i = 1; $i -le $limit; $i++) {
        $item = $items.Item($i)
        if ($null -ne $item) {
            if ($item.SenderName -like "*김하영*" -or $item.Subject -like "*AI 과제*") {
                $results += [PSCustomObject]@{
                    Index = $i
                    Sender = $item.SenderName
                    Subject = $item.Subject
                    Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                }
            }
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}