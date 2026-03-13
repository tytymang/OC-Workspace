
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $sent = $namespace.GetDefaultFolder(5)
    
    # 1. 보낸 편지함에서 '가장 최신 메일' 무조건 가져오기
    $items = $sent.Items
    $items.Sort("[SentOn]", $true)
    $latestMail = $items.Item(1)

    Write-Host "Found Mail: $($latestMail.Subject)"

    # 2. 수신인(To) 추출
    $recipients = @()
    foreach ($recip in $latestMail.Recipients) {
        if ($recip.Type -eq 1) { # olTo
            $recipients += $recip.Address
        }
    }

    if ($recipients.Count -eq 0) {
        throw "수신인이 없습니다."
    }

    # 3. 일정 생성 (유니코드 조립 - '2월 마감 회의 확인')
    # 2월:[char]50+[char]50900;  마감:[char]32+[char]47560+[char]44048;  회의:[char]32+[char]54924+[char]51032;  확인:[char]32+[char]54869+[char]51064;
    $subject = [string][char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]54869 + [char]51064;
    
    $appt = $outlook.CreateItem(1) # olAppointmentItem
    $appt.Subject = $subject
    $appt.Start = (Get-Date).AddHours(1).ToString("yyyy-MM-dd HH:00:00")
    $appt.Duration = 60
    $appt.MeetingStatus = 1 # olMeeting
    
    foreach ($addr in $recipients) {
        $appt.Recipients.Add($addr)
    }

    # 4. 화면에 띄우기
    $appt.Display()
    
    "SUCCESS: Created meeting based on latest mail: " + $latestMail.Subject
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\error_latest.log"
    throw $_
}
