$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $found = $null
    for ($i = 1; $i -le 50; $i++) {
        $item = $items.Item($i)
        if ($item.SenderName -like "*이수정*") {
            $found = $item
            break
        }
    }
    
    if ($null -eq $found) {
        Write-Output "NOT_FOUND"
        exit
    }

    $results = [PSCustomObject]@{
        Sender = $found.SenderName
        Subject = $found.Subject
        Received = $found.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        Body = $found.Body
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}