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
$senders = @("SSCBPM", "NGUYEN THI THUAN(NGUYEN, THITHUAN/SENIOR ASSOCIATE)")

foreach ($s in $senders) {
    foreach ($vip in $vipNames) {
        if ($s -like "*$vip*") {
            Write-Output "Matched: $s with $vip"
        }
    }
}
