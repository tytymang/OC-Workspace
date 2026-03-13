
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    # 최근 10개 메일 중 한글 깨진 것들 식별 및 발신자 확인
    for ($i = 1; $i -le 10; $i++) {
        $item = $items.Item($i)
        $results += [PSCustomObject]@{
            Index = $i
            Subject = $item.Subject
            Sender = $item.SenderName
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output $_.Exception.Message
}
