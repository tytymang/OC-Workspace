
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$inbox = $namespace.GetDefaultFolder(6)

$startDate = Get-Date -Year 2024 -Month 12 -Day 1
$endDate = Get-Date -Year 2025 -Month 11 -Day 1

Write-Host "Searching Zionex related items..."

$items = $inbox.Items
$total = $items.Count
$foundCount = 0

for ($i = $total; $i -gt 0; $i--) {
    try {
        $item = $items.Item($i)
        $received = $item.ReceivedTime
        if ($received -ge $startDate -and $received -lt $endDate) {
            $s = $item.Subject
            $b = $item.Body
            # 영어 키워드로 먼저 필터링 (한글 깨짐 방지)
            if ($s.ToLower().Contains("zionex") -or $b.ToLower().Contains("zionex")) {
                Write-Host "DATE_MSG: $($received.ToString('yyyy-MM-dd'))"
                Write-Host "SUB_MSG: $s"
                $foundCount++
            }
        }
        # 너무 오래 전으로 가면 중단 (최적화)
        if ($received -lt $startDate.AddMonths(-1)) { break }
    } catch {}
}
Write-Host "Total found: $foundCount"
