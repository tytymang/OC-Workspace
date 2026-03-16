
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    function Search-Folder($folder) {
        $items = $folder.Items
        # Search for digits pattern in subject (Trip.com orders usually have numbers)
        $found = $items | Where-Object { $_.Subject -match "\d{10,}" -or $_.Body -match "Trip.com" }
        foreach ($m in $found) {
            Write-Host "---"
            Write-Host "Subject: $($m.Subject)"
            Write-Host "Received: $($m.ReceivedTime)"
            if ($m.Subject -match "Trip.com") {
                Write-Host "Body: $($m.Body.Substring(0, [Math]::Min(2000, $m.Body.Length)))"
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
