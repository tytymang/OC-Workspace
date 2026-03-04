
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$sentItems = $namespace.GetDefaultFolder(5)

# 인덱스로 직접 접근하여 3번째 메일(02-25 16:59) 내용 확인
$items = $sentItems.Items
$total = $items.Count
$item = $items.Item($total - 2) # 최근 10개 중 3번째

Write-Host "---MAIL_START---"
Write-Host "Subject: $($item.Subject)"
Write-Host "SentOn: $($item.SentOn)"
Write-Host "To: $($item.To)"
Write-Host "Body: $($item.Body)"
Write-Host "---MAIL_END---"
