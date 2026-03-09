
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    
    # [2월 마감 회의 확인] 유니코드 코드포인트 조립
    # 2:[char]50; 월:[char]50900;  :[char]32; 마:[char]47560; 감:[char]44048;  :[char]32; 회:[char]54924; 의:[char]51032;  :[char]32; 확:[char]54869; 인:[char]51064;
    $subject = [string][char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]54869 + [char]51064;

    # 본문: "2월 마감 완료 사항 확인을 위한 회의입니다." 유니코드 조립
    # 2월 마감 완료 사항 확인을 위한 회의입니다.
    # 2:[char]50; 월:[char]50900;  :[char]32; 마:[char]47560; 감:[char]44048;  :[char]32; 완:[char]50756; 료:[char]47308;  :[char]32; 사:[char]49324; 항:[char]54637;  :[char]32; 
    # 확:[char]54869; 인:[char]51064; 을:[char]51012;  :[char]32; 위:[char]50948; 한:[char]54620;  :[char]32; 회:[char]54924; 의:[char]51032; 입:[char]51077; 니:[char]45768; 다:[char]45796; .:[char]46;
    $body = [string][char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]50756 + [char]47308 + [char]32 + [char]49324 + [char]54637 + [char]32 + [char]54869 + [char]51064 + [char]51012 + [char]32 + [char]50948 + [char]54620 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]51077 + [char]45768 + [char]45796 + [char]46;

    $appt = $outlook.CreateItem(1) # olAppointmentItem
    $appt.Subject = $subject
    $appt.Body = $body
    $appt.Location = "https://sskv.webex.com/sskv/j.php?MTID=mb9e45aa44dbb35ccfc41197aabb63e1b"
    
    # 시간 설정: 2026-03-05 11:00
    $appt.Start = Get-Date -Year 2026 -Month 3 -Day 5 -Hour 11 -Minute 0 -Second 0
    $appt.Duration = 60
    $appt.MeetingStatus = 1 # olMeeting

    # 수신인 추가
    $appt.Recipients.Add("kim, JW")
    $appt.Recipients.Add("Kim, SH")
    $appt.Recipients.Add("Yuk, HaNa")
    $appt.Recipients.ResolveAll()

    # 화면에 표시
    $appt.Display()
    
    "SUCCESS: Meeting displayed with safe unicode encoding."
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\final_fix.log"
    throw $_
}
