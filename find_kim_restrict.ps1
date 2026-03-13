$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # Restrict 사용
    $filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0042001f"" LIKE '%Kim, HaYoung%'"
    $items = $inbox.Items.Restrict($filter)
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    $count = [Math]::Min($items.Count, 10)
    for ($i = 1; $i -le $count; $i++) {
        $item = $items.Item($i)
        $results += [PSCustomObject]@{
            Index = $i
            Sender = $item.SenderName
            Subject = $item.Subject
            Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}