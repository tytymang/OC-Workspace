$ErrorActionPreference = "Stop"
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$results = @{ emails = @(); calendar = @() }
$now = Get-Date

$v1 = -join @([char]0xC774, [char]0xC815, [char]0xD6C8)
$v2 = -join @([char]0xC774, [char]0xC815, [char]0xC6B0)
$v3 = -join @([char]0xC774, [char]0xC0C1, [char]0xBB34)
$v4 = -join @([char]0xC774, [char]0xC601, [char]0xC91C)
$vips = @($v1, $v2, $v3, $v4)

function Check-Folder($f) {
    try {
        $items = $f.Items.Restrict("[UnRead] = true AND [ReceivedTime] > '2026-03-25 13:16'")
        foreach ($i in $items) {
            foreach ($v in $vips) {
                if ($i.SenderName -like "*$v*") {
                    $results.emails += @{ Sender = $i.SenderName; Subject = $i.Subject }
                    break
                }
            }
        }
    } catch {}
    if ($f.Name -match "Inbox" -or $f.Name -match "받은") {
        foreach ($s in $f.Folders) { Check-Folder $s }
    }
}

foreach ($r in $namespace.Folders) {
    try {
        $inbox = $r.Folders.Item("Inbox")
        if ($null -ne $inbox) { Check-Folder $inbox }
    } catch {
        try {
            $kb = -join @([char]0xBC1B, [char]0xC740, [char]0x0020, [char]0xD3B8, [char]0xC9C0, [char]0xD568)
            $kr = $r.Folders | Where-Object { $_.Name -match $kb }
            if ($null -ne $kr) { Check-Folder $kr }
        } catch {}
    }
}

try {
    $cal = $namespace.GetDefaultFolder(9)
    $items = $cal.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
    $filter = "[Start] >= '$($now.ToString("yyyy-MM-dd HH:mm"))' AND [Start] <= '$($now.AddHours(2).ToString("yyyy-MM-dd HH:mm"))'"
    foreach ($item in $cal.Items.Restrict($filter)) {
        $results.calendar += @{ Subject = $item.Subject; Start = $item.Start.ToString("HH:mm") }
    }
} catch {}

$results | ConvertTo-Json -Compress
