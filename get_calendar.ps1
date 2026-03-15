
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
    $items = $calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
    
    $today = Get-Date -Hour 0 -Minute 0 -Second 0
    $tomorrow = $today.AddDays(1)
    
    foreach ($item in $items) {
        if ($item.Start -ge $today -and $item.Start -lt $tomorrow) {
            Write-Output "CAL: [$($item.Start.ToString('HH:mm'))] $($item.Subject)"
        }
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
