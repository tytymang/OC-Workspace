
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    
    # Try searching for "Confirmation" or "Booking"
    $found = $items | Where-Object { $_.Subject -match "Confirmation" -or $_.Subject -match "Booking" } | Select-Object -First 10
    
    foreach ($m in $found) {
        Write-Host "---"
        Write-Host "Subject: $($m.Subject)"
        Write-Host "Received: $($m.ReceivedTime)"
        Write-Host "Body: $($m.Body.Substring(0, [Math]::Min(1000, $m.Body.Length)))"
    }
} catch {
    Write-Error $_.Exception.Message
}
