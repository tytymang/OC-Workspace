
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    function Search-Folder($folder) {
        $items = $folder.Items
        # Filter for items with 'Trip.com' in subject or body
        $found = $items | Where-Object { $_.Subject -match "Trip" -or $_.Body -match "Trip" }
        foreach ($m in $found) {
            Write-Host "---"
            Write-Host "Folder: $($folder.FolderPath)"
            Write-Host "Subject: $($m.Subject)"
            Write-Host "Received: $($m.ReceivedTime)"
            $bodySnippet = $m.Body.Substring(0, [Math]::Min(1000, $m.Body.Length))
            Write-Host "Body: $bodySnippet"
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
