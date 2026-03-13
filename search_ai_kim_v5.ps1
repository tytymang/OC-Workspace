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
    
    for ($i = 1; $i -le 100; $i++) {
        $item = $items.Item($i)
        if ($null -eq $item) { break }
        if ($item.Subject -like "*AI*") {
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