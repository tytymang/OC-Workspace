
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $sentFolder = $namespace.GetDefaultFolder(5)
    
    # 최신 메일 수신인 다시 확보
    $latestMail = $sentFolder.Items | Sort-Object SentOn -Descending | Select-Object -First 1
    $recipients = @()
    foreach ($recip in $latestMail.Recipients) {
        if ($recip.Type -eq 1) {
            $recipients += $recip.Address
        }
    }

    # '2월 마감 회의 확인' 유니코드 조립
    $subject = [char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]54869 + [char]51064
    
    # 일정 생성
    $appt = $outlook.CreateItem(1)
    $appt.Subject = $subject
    $appt.Location = "https://sskv.webex.com/sskv/j.php?MTID=mb9e45aa44dbb35ccfc41197aabb63e1b"
    $appt.Body = "2월 마감 완료 사항 확인을 위한 회의입니다. 일정 확인 부탁드립니다."
    
    # 시간 설정: 오늘 오전 11시 (2026-03-05 11:00)
    $appt.Start = Get-Date -Year 2026 -Month 3 -Day 5 -Hour 11 -Minute 0 -Second 0
    $appt.Duration = 60
    $appt.MeetingStatus = 1
    
    foreach ($addr in $recipients) {
        $appt.Recipients.Add($addr)
    }

    # 화면에 표시
    $appt.Display()
    
    "SUCCESS: Displayed full meeting details for " + ($recipients -join ", ")
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\meeting_retry.log"
    throw $_
}
