$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$addressBook = $namespace.AddressLists.Item("Global Address List")

$names = @("이정훈", "이상무", "이영주")
$results = @()

foreach ($name in $names) {
    try {
        $entry = $addressBook.AddressEntries.Item($name)
        if ($entry) {
            $user = $entry.GetExchangeUser()
            if ($user) {
                $results += @{ Name = $name; Email = $user.PrimarySmtpAddress }
            } else {
                $results += @{ Name = $name; Email = $entry.Address }
            }
        }
    } catch {
        # Search via Name
        $results += @{ Name = $name; Email = "Not Found in GAL" }
    }
}

$results | ConvertTo-Json | Out-File -FilePath "vip_search.json" -Encoding UTF8
