$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $secretaryFolder = $inbox.Folders.Item(7)
    $kimFolder = $secretaryFolder.Folders.Item(2)

    $items = $kimFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    # 상위 10개 메일 제목 출력
    $results = @()
    $count = [Math]::Min(10, $items.Count)
    for ($i = 1; $i -le $count; $i++) {
        $results += $items.Item($i).Subject
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}