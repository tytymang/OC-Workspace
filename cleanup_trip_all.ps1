
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
    $items = $calendar.Items
    
    # Delete all appointments on 3/18 and 4/2 that match "VN"
    $today = Get-Date -Format "yyyy-MM-dd"
    $targetDates = @("2026-03-18", "2026-04-02", "2026-04-03")
    
    foreach ($item in $items) {
        $startStr = $item.Start.ToString("yyyy-MM-dd")
        if ($targetDates -contains $startStr) {
             # If it looks like a business trip item
             if ($item.Subject -match "VN" -or $item.Body -match "VN") {
                Write-Host "DELETING: $($item.Subject)"
                $item.Delete()
             }
        }
    }
} catch {
    Write-Error $_.Exception.Message
}
