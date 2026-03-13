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
    $count = [Math]::Min(100, $items.Count)
    
    for ($i = 1; $i -le $count; $i++) {
        $item = $items.Item($i)
        $subj = $item.Subject
        if ($subj -like "*AI*" -and $subj -like "*과제*") {
            $results += [PSCustomObject]@{
                Index = $i
                Subject = $subj
            }
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}