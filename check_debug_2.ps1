$v1 = -join @([char]51060, [char]51221, [char]54984) # 이정훈
$v2 = -join @([char]51060, [char]49345, [char]47924) # 이상무
$v3 = -join @([char]51060, [char]50689, [char]51452) # 이영주
$v4 = -join @([char]51060, [char]51221, [char]50864) # 이정우
$v5 = -join @([char]44608, [char]49457, [char]51452) # 김성주
$v6 = -join @([char]51060, [char]49457, [char]51008) # 이성은
$v7 = -join @([char]44608, [char]54616, [char]50689) # 김하영
$v8 = -join @([char]52572, [char]48337, [char]49692) # 최병순
$v9 = -join @([char]51060, [char]46041, [char]49437) # 이동석

$vipNames = @($v1, $v2, $v3, $v4, $v5, $v6, $v7, $v8, $v9)

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $unreadItems = $inbox.Items.Restrict("[UnRead] = true")
    $unreadItems.Sort("[ReceivedTime]", $true)

    $results = @()
    $timeLimit = (Get-Date).AddHours(-1)
    foreach ($item in $unreadItems) {
        if ($item.ReceivedTime -gt $timeLimit) {
            $isVip = $false
            foreach ($vip in $vipNames) {
                if (-not [string]::IsNullOrEmpty($vip) -and $item.SenderName -match $vip) {
                    $isVip = $true
                    break
                }
            }
            if ($isVip) {
                $results += @{
                    ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                    Sender = $item.SenderName
                    Subject = $item.Subject
                }
            }
        }
    }
    $results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\hb_vip_mails_debug.json" -Encoding UTF8
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\hb_mail_err_debug.txt"
}
