$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)
$items = $inbox.Items
$items.Sort("[ReceivedTime]", $true)

$results = @()
for ($i = 1; $i -le 500 -and $i -le $items.Count; $i++) {
    $item = $items.Item($i)
    $results += @{ Name = $item.SenderName; Email = $item.SenderEmailAddress }
}

$results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\inbox_deep_scan.json" -Encoding UTF8
