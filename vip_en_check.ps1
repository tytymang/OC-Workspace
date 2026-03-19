$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$gal = $namespace.GetGlobalAddressList()

$results = @()
$n1 = -join @([char]51060, [char]51221, [char]54984) # 이정훈
$n2 = -join @([char]51060, [char]49345, [char]47924) # 이상무
$n3 = -join @([char]51060, [char]50689, [char]51452) # 이영주
$targets = @($n1, $n2, $n3)

# Search specifically for Lee, JungHoon etc. in English if Korean fails
$targets_en = @("Lee, JungHoon", "Lee, SangMoo", "Lee, YoungJu")

foreach ($name in $targets_en) {
    try {
        $entry = $gal.AddressEntries.Item($name)
        if ($entry) {
            $user = $entry.GetExchangeUser()
            $email = if ($user) { $user.PrimarySmtpAddress } else { $entry.Address }
            $results += @{ Name = $entry.Name; Email = $email }
        }
    } catch { }
}

$results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\vip_en_check.json" -Encoding UTF8
