
try {
    $outlook = New-Object -ComObject Outlook.Application
    $calendar = $outlook.GetNamespace("MAPI").GetDefaultFolder(9)
    $now = Get-Date
    $twoHours = $now.AddHours(2)
    $items = $calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
    foreach($item in $items) {
        if ($item.Start -ge $now -and $item.Start -le $twoHours) {
            Write-Output "UPCOMING: $($item.Start) - $($item.Subject)"
        }
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
