
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")

    # Look for items starting from today or recently added
    $today = Get-Date -Format "yyyy-MM-dd"
    $filter = "[Start] >= '$today 00:00 AM'"
    $restrictedItems = $items.Restrict($filter)

    foreach ($item in $restrictedItems) {
        Write-Host "---"
        Write-Host "Subject: $($item.Subject)"
        Write-Host "Start: $($item.Start)"
        Write-Host "End: $($item.End)"
        Write-Host "Body: $($item.Body.Substring(0, [Math]::Min(200, $item.Body.Length)))"
    }
} catch {
    Write-Error $_.Exception.Message
}
