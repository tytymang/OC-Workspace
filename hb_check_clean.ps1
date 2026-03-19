try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    $inbox = $namespace.GetDefaultFolder(6)
    $unreadItems = $inbox.Items.Restrict("[UnRead] = true")
    $unreadItems.Sort("[ReceivedTime]", $true)

    $v1 = -join @([char]51060, [char]51221, [char]54984)
    $v2 = -join @([char]51060, [char]49345, [char]47924)
    $v3 = -join @([char]51060, [char]50689, [char]51452)
    $v4 = -join @([char]51060, [char]51221, [char]50864)
    $v5 = -join @([char]44608, [char]49457, [char]51452)
    $v6 = -join @([char]51060, [char]49457, [char]51008)
    $v7 = -join @([char]44608, [char]54616, [char]50689)
    $v8 = -join @([char]52572, [char]48337, [char]49692)
    $v9 = -join @([char]51060, [char]46041, [char]49437)

    $vipNames = @($v1, $v2, $v3, $v4, $v5, $v6, $v7, $v8, $v9)

    $mails = @()
    foreach ($item in $unreadItems) {
        if ($item.ReceivedTime -gt (Get-Date).AddMinutes(-45)) {
            $isVip = $false
            foreach ($vip in $vipNames) {
                if ($item.SenderName -like "*$vip*") {
                    $isVip = $true
                    break
                }
            }
            if ($isVip) {
                $mails += @{
                    ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                    Sender = $item.SenderName
                    Subject = $item.Subject
                }
            }
        }
    }

    $calendar = $namespace.GetDefaultFolder(9)
    $startTime = Get-Date
    $endTime = $startTime.AddHours(2)
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    $filter = "[Start] >= '" + $startTime.ToString("g") + "' AND [Start] <= '" + $endTime.ToString("g") + "'"
    $recentApps = $items.Restrict($filter)

    $cals = @()
    foreach ($app in $recentApps) {
        $cals += @{
            Subject = $app.Subject
            Start = $app.Start.ToString("yyyy-MM-dd HH:mm")
        }
    }

    $res = @{ Mails = $mails; Calendar = $cals }
    $res | ConvertTo-Json -Depth 5 | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\hb_1347_clean.json" -Encoding UTF8
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\hb_err_1347_clean.txt"
}
