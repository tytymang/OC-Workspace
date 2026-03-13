$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
    $now = Get-Date
    $later = $now.AddHours(2)
    
    $filter = "[Start] >= '$($now.ToString("g"))' AND [Start] <= '$($later.ToString("g"))'"
    $items = $calendar.Items.Restrict($filter)
    
    $results = @()
    foreach ($item in $items) {
        $results += [PSCustomObject]@{
            Subject = $item.Subject
            Start = $item.Start.ToString("yyyy-MM-dd HH:mm")
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}