
$ErrorActionPreference = "Stop"
function Search-Folders($folder) {
    $items = $folder.Items
    $found = $items | Where-Object { $_.Subject -match "Trip.com" -or $_.Body -match "Trip.com" }
    foreach ($item in $found) {
        Write-Host "---"
        Write-Host "Folder: $($folder.Name)"
        Write-Host "Subject: $($item.Subject)"
        Write-Host "Received: $($item.ReceivedTime)"
        $body = $item.Body
        Write-Host "Body: $($body.Substring(0, [Math]::Min(1000, $body.Length)))"
    }
    foreach ($subFolder in $folder.Folders) {
        Search-Folders $subFolder
    }
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    foreach ($store in $namespace.Stores) {
        Search-Folders $store.GetRootFolder()
    }
} catch {
    Write-Error $_.Exception.Message
}
