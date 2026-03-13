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
    
    # 200개까지 검색 범위를 넓히고 'AI'만 포함해도 리스팅
    for ($i = 1; $i -le [Math]::Min(200, $items.Count); $i++) {
        $item = $items.Item($i)
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