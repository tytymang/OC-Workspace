[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$outlook = New-Object -ComObject Outlook.Application
$app = $outlook.CreateItem(1)
$app.Subject = "SCM 구축 협의"
$app.Start = "2026-02-27 15:00"
$app.End = "2026-02-27 16:00"
$app.Location = "1층 108 회의실 (https://sskv.webex.com/sskv/j.php?MTID=m35ed789b69ac6c98c2496b844b12571f)"
$app.Body = @"
5. 협의 내용
- 단계별 SCM 구축 전략
- 1단계 과제 상세 협의 > T3 Smart SCM Platform Migration (기능 개선, 이관 전략 等)
"@

$attendees = @("최현구", "이동석", "김제림", "김종원", "강동민", "육하나", "김승현", "정혁곤")
foreach ($person in $attendees) {
    $recipient = $app.Recipients.Add($person)
    $recipient.Type = 1
}

$app.Save()
Write-Output "Done: $($app.Subject)"
