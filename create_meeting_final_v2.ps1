
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $sentFolder = $namespace.GetDefaultFolder(5)
    
    # 1. 보낸 메일함에서 가장 최신 메일로부터 수신인(To) 주소 직접 확보
    $latestMail = $sentFolder.Items | Sort-Object SentOn -Descending | Select-Object -First 1
    if ($latestMail -eq $null) { throw "보낸 메일을 찾을 수 없습니다." }
    
    $recipientsAddresses = @()
    foreach ($recip in $latestMail.Recipients) {
        if ($recip.Type -eq 1) { # olTo
            $recipientsAddresses += $recip.Address
        }
    }

    # 2. 유니코드 코드포인트 조립 (한글 절대 사수)
    # 제목: [2월 마감 회의 확인]
    $subject = [string][char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]54869 + [char]51064
    
    # 본문: [2월 마감 완료 사항 확인을 위한 회의입니다.]
    $body = [string][char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]50756 + [char]47308 + [char]32 + [char]49324 + [char]54637 + [char]32 + [char]54869 + [char]51064 + [char]51012 + [char]32 + [char]50948 + [char]54620 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]51077 + [char]45768 + [char]45796 + [char]46
    
    # 3. 일정(모임) 생성 및 데이터 주입
    $appt = $outlook.CreateItem(1) # olAppointmentItem
    $appt.Subject = $subject
    $appt.Body = $body
    $appt.Location = "https://sskv.webex.com/sskv/j.php?MTID=mb9e45aa44dbb35ccfc41197aabb63e1b"
    $appt.Start = Get-Date -Year 2026 -Month 3 -Day 5 -Hour 11 -Minute 0 -Second 0
    $appt.Duration = 60
    $appt.MeetingStatus = 1 # olMeeting (필수 참석자 활성화)

    # 4. 수신인 추가 및 확인(Resolve)
    foreach ($addr in $recipientsAddresses) {
        $newRecip = $appt.Recipients.Add($addr)
        $newRecip.Type = 1 # olRequired
    }
    $appt.Recipients.ResolveAll()

    # 5. 화면에 즉시 표시
    $appt.Display()
    
    "SUCCESS: Meeting created with full details and resolved recipients."
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\final_attempt_meeting.log"
    throw $_
}
