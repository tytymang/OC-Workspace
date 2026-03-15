
try {
    $outlook = New-Object -ComObject Outlook.Application
    $calendar = $outlook.GetNamespace("MAPI").GetDefaultFolder(9)
    $today = Get-Date -Hour 0 -Minute 0 -Second 0
    $tomorrow = $today.AddDays(1)
    $items = $calendar.Items
    $items.Restrict("[Start] >= '$($today.ToString('g'))' AND [Start] < '$($tomorrow.ToString('g'))'") | ForEach-Object {
        Write-Output "CAL: $($_.Start) - $($_.Subject)"
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
