
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    function Search-Folder($folder) {
        $items = $folder.Items
        $found = $items | Where-Object { $_.Subject -match "Flight" -or $_.Subject -match "Ticket" -or $_.Subject -match "Booking" -or $_.Subject -match "Trip" }
        foreach ($m in $found) {
            Write-Host "---"
            Write-Host "Subject: $($m.Subject)"
            Write-Host "Received: $($m.ReceivedTime)"
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
