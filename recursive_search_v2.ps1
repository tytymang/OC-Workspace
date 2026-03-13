$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $results = @()
    
    function SearchFolder($folder) {
        $found = @()
        try {
            foreach ($item in $folder.Items) {
                if ($item.SenderName -like "*김하영*" -and $item.Subject -like "*AI 과제*") {
                    $found += [PSCustomObject]@{
                        Sender = $item.SenderName
                        Subject = $item.Subject
                        Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                        FolderName = $folder.Name
                    }
                }
            }
            foreach ($sub in $folder.Folders) {
                $found += SearchFolder($sub)
            }
        } catch {}
        return $found
    }

    $allResults = SearchFolder($inbox)
    $allResults | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}