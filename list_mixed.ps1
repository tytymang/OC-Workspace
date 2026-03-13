$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    for ($i = 1; $i -le 20; $i++) {
        $item = $items.Item($i)
        if ($null -ne $item) {
            $results += [PSCustomObject]@{
                Index = $i
                Sender = $item.SenderName
                Subject = $item.Subject
                Unread = $item.UnRead
            }
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}