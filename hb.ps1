$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$in = $ns.GetDefaultFolder(6)
$ca = $ns.GetDefaultFolder(9)
$now = Get-Date
$later = $now.AddHours(2)
$recent = $now.AddMinutes(-30)

"EMAILS:"
$in.Items | ? { $_.ReceivedTime -ge $recent } | % { $_.SenderName + " | " + $_.Subject }

"CALENDAR:"
$ca.Items | ? { $_.Start -ge $now -and $_.Start -le $later } | % { $_.Start.ToString("HH:mm") + " | " + $_.Subject }

"GIT:"
git status --short
