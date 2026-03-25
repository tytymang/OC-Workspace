$ErrorActionPreference = "Stop"
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$now = Get-Date
$lastReportTime = [DateTime]"2026-03-25 08:10:00"

# VIPs as Hex/Unicode
$lee_j_h = -join @([char]0xC774, [char]0xC815, [char]0xD6C8) # 이정훈
$lee_j_w = -join @([char]0xC774, [char]0xC815, [char]0xC6B0) # 이정우
$lee_s_m = -join @([char]0xC774, [char]0xC131, [char]0xBB34) # 이상무 (Wait, 상 is C131 or C121?)
$lee_y_j = -join @([char]0xC774, [char]0xC601, [char]0xC9FC) # 이영주

# Let's just use part of the name or check the SenderName more carefully
$results = @()
function Scan-Folders($folder) {
    try {
        $filter = "[UnRead] = true AND [SentOn] > '$($lastReportTime.ToString("g"))'"
        $items = $folder.Items.Restrict($filter)
        foreach ($item in $items) {
            $sender = $item.SenderName
            $results += [PSCustomObject]@{
                Sender = $sender
                Subject = $item.Subject
                SentTime = $item.SentOn.ToString("HH:mm")
            }
        }
    } catch {}
    foreach ($sub in $folder.Folders) { Scan-Folders $sub }
}

foreach ($root in $namespace.Folders) {
    Scan-Folders $root
}

$results | ConvertTo-Json -Compress
