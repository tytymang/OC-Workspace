
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $calendar = $ns.GetDefaultFolder(9)
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    $results = @()
    $count = 0
    foreach($item in $items) {
        if ($item.Start -ge (Get-Date).Date) {
            $results += [PSCustomObject]@{
                Start   = $item.Start.ToString("yyyy-MM-dd HH:mm")
                Subject = $item.Subject
            }
            $count++
        }
        if ($count -ge 5) { break }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output $_.Exception.Message
}
