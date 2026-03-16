
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    function Search-Folder($folder) {
        if ($folder.Name -match "Trip") {
            Write-Host "FOUND FOLDER: $($folder.FolderPath)"
            foreach ($item in $folder.Items) {
                Write-Host "---"
                Write-Host "Subject: $($item.Subject)"
                Write-Host "Received: $($item.ReceivedTime)"
                Write-Host "Body: $($item.Body.Substring(0, [Math]::Min(1000, $item.Body.Length)))"
            }
        }
        foreach ($sub in $folder.Folders) {
            Search-Folder $sub
        }
    }
    
    foreach ($store in $namespace.Stores) {
        Search-Folder $store.GetRootFolder()
    }
} catch {
    Write-Error $_.Exception.Message
}
