$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Calendar = $Namespace.GetDefaultFolder(9)

$todayStr = (Get-Date).ToString("yyyy-MM-dd")
$Start = Get-Date "$todayStr 00:00:00"
$End = Get-Date "$todayStr 23:59:59"

$Filter = "[Start] >= '" + $Start.ToString("g") + "' AND [Start] <= '" + $End.ToString("g") + "'"
$Items = $Calendar.Items
$Items.IncludeRecurrences = $true
$Items.Sort("[Start]")
$Res = $Items.Restrict($Filter)

$out = @()
foreach ($item in $Res) {
    $out += [PSCustomObject]@{
        Subject = $item.Subject
        Start = $item.Start.ToString("HH:mm")
        End = $item.End.ToString("HH:mm")
        Location = $item.Location
    }
}
$out | ConvertTo-Json
