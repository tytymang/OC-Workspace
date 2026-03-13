$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    function Search-Folder($folder) {
        $results = @()
        # 해당 폴더에서 찾기
        foreach ($item in $folder.Items) {
            if ($item.SenderName -like "*김하영*" -and $item.Subject -like "*AI 과제*") {
                $results += [PSCustomObject]@{
                    Sender = $item.SenderName
                    Subject = $item.Subject
                    Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                    FolderName = $folder.Name
                }
            }
        }
        # 하위 폴더 검색
        foreach ($sub in $folder.Folders) {
            $results += Search-Folder($sub)
        }
        return $results
    }

    $allResults = Search-Folder($inbox)
    $allResults | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}