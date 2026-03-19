$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$results = @()
# List using code points for robustness
$n1 = -join @([char]51060, [char]51221, [char]54984) # 이정훈
$n2 = -join @([char]51060, [char]49345, [char]47924) # 이상무
$n3 = -join @([char]51060, [char]50689, [char]51452) # 이영주
$n4 = -join @([char]51060, [char]51221, [char]50864) # 이정우
$n5 = -join @([char]44608, [char]49457, [char]51452) # 김성주
$n6 = -join @([char]51060, [char]49457, [char]51008) # 이성은
$n7 = -join @([char]44608, [char]54616, [char]50689) # 김하영
$n8 = -join @([char]52572, [char]48337, [char]49692) # 최병순
$n9 = -join @([char]51060, [char]46041, [char]49437) # 이동석

$targets = @($n1, $n2, $n3, $n4, $n5, $n6, $n7, $n8, $n9)

foreach ($name in $targets) {
    try {
        $recip = $namespace.CreateRecipient($name)
        if ($recip.Resolve()) {
            $ae = $recip.AddressEntry
            $user = $ae.GetExchangeUser()
            $results += @{
                Name = $ae.Name
                Email = if ($user) { $user.PrimarySmtpAddress } else { $ae.Address }
            }
        }
    } catch { }
}

$results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\vip_config.json" -Encoding UTF8
