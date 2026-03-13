$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # "이상무"가 들어가는 폴더 찾기
    $targetFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -like "*이상무*") {
            $targetFolder = $f
            break
        }
    }

    if ($null -eq $targetFolder) {
        Write-Output "FOLDER_NOT_FOUND"
        exit
    }
    
    $items = $targetFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    $limit = if ($items.Count -lt 50) { $items.Count } else { 50 }
    for ($i = 1; $i -le $limit; $i++) {
        $item = $items.Item($i)
        if ($item.SenderName -like "*김하영*") {
             $results += [PSCustomObject]@{
                Sender = $item.SenderName
                Subject = $item.Subject
                Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            }
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}