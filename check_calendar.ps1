$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9) # olFolderCalendar
    
    $start = Get-Date
    $end = (Get-Date).AddHours(2)
    
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    
    $filter = "[Start] >= '$($start.ToString("yyyy-MM-dd HH:mm"))' AND [Start] <= '$($end.ToString("yyyy-MM-dd HH:mm"))'"
    $restrictedItems = $items.Restrict($filter)
    
    $results = @()
    foreach ($item in $restrictedItems) {
        $results += [PSCustomObject]@{
            Subject = $item.Subject
            Start = $item.Start.ToString("yyyy-MM-dd HH:mm")
            Location = $item.Location
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}