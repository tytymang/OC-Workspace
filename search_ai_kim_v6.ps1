$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $secretaryFolder = $inbox.Folders.Item(7)
    $kimFolder = $secretaryFolder.Folders.Item(2)
    
    $results = @()
    $items = $kimFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $count = $items.Count
    if ($count -gt 100) { $count = 100 }
    
    for ($i = 1; $i -le $count; $i++) {
        $item = $items.Item($i)
        if ($item.Subject -like "*AI*" -and $item.Subject -like "*과제*") {
            $results += [PSCustomObject]@{
                Index = $i
                Subject = $item.Subject
            }
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}