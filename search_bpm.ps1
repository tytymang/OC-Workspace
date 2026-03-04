
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$inbox = $namespace.GetDefaultFolder(6)

$items = $inbox.Items
$total = $items.Count
Write-Host "Total items in Inbox: $total"

# 최근 100개만 무조건 출력
for ($i = $total; $i -gt ($total - 100); $i--) {
    try {
        $item = $items.Item($i)
        Write-Host "[$($item.ReceivedTime.ToString('yyyy-MM-dd'))] $($item.Subject)"
    } catch {}
}
