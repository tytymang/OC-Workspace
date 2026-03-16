
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
    
    # Subject contains broken characters based on the image
    $items = $calendar.Items
    $found = $items | Where-Object { $_.Subject -match "VN" -and ($_.Subject -match "異" -or $_.Subject -match "쒓") }
    
    foreach ($m in $found) {
        Write-Host "DELETING: $($m.Subject)"
        $m.Delete()
    }
} catch {
    Write-Error $_.Exception.Message
}
