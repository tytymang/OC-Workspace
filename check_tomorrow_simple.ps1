
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $calendar = $ns.GetDefaultFolder(9)
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    
    $tomorrow = (Get-Date).AddDays(1).Date
    $results = @()
    foreach($item in $items) {
        if ($item.Start.Date -eq $tomorrow) {
            $results += [PSCustomObject]@{
                Start   = $item.Start.ToString("HH:mm")
                Subject = $item.Subject
            }
        }
        if ($item.Start.Date -gt $tomorrow) { break }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output $_.Exception.Message
}
