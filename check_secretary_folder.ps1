$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $folder = $inbox.Folders.Item("비서/기획")
    
    $items = $folder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    $limit = if ($items.Count -lt 30) { $items.Count } else { 30 }
    for ($i = 1; $i -le $limit; $i++) {
        $item = $items.Item($i)
        $results += [PSCustomObject]@{
            Sender = $item.SenderName
            Subject = $item.Subject
            Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}