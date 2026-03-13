
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $sent = $namespace.GetDefaultFolder(5)
    
    # 1. '마감'이 포함된 최신 보낸 메일 찾기 (최근 200건)
    $foundMail = $null
    $items = $sent.Items
    $items.Sort("[SentOn]", $true)
    
    for ($i = 1; $i -le [Math]::Min($items.Count, 200); $i++) {
        $item = $items.Item($i)
        if ($item.Subject -like "*마감*") {
            $foundMail = $item
            break
        }
    }

    if ($foundMail -eq $null) {
        throw "제목에 '마감'이 포함된 메일을 찾을 수 없습니다."
    }

    # 2. 수신인(To) 추출
    $recipients = @()
    foreach ($recip in $foundMail.Recipients) {
        if ($recip.Type -eq 1) { # olTo
            $recipients += $recip.Address
        }
    }

    if ($recipients.Count -eq 0) {
        throw "수신인이 없습니다."
    }

    # 3. 일정 생성 (유니코드 조립 방식 - '2월 마감 회의')
    # 2:[char]50; 월:[char]50900;  :[char]32; 마:[char]47560; 감:[char]44048;  :[char]32; 회:[char]54924; 의:[char]51032;
    $subject = [string][char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]54924 + [char]51032;
    
    $appt = $outlook.CreateItem(1) # 1 = olAppointmentItem
    $appt.Subject = $subject
    $appt.Start = (Get-Date).AddHours(1) # 현재 시간 1시간 후로 임시 설정
    $appt.Duration = 60
    $appt.MeetingStatus = 1 # 1 = olMeeting (모임으로 설정)
    
    foreach ($addr in $recipients) {
        $appt.Recipients.Add($addr)
    }

    # 4. 화면에 띄우기 (주인님 확인용)
    $appt.Display()
    
    "SUCCESS: Created meeting for " + ($recipients -join ", ")
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\error_meeting.log"
    throw $_
}
