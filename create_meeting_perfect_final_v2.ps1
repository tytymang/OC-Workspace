
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $sentFolder = $namespace.GetDefaultFolder(5)
    
    # 09:26:41에 발송된 메일 직접 타겟팅 (RE: 2월 마감 완료 확인)
    $items = $sentFolder.Items
    $items.Sort("[SentOn]", $true)
    
    $targetMail = $null
    for ($i = 1; $i -le 30; $i++) {
        $mail = $items.Item($i)
        # 제목에 '2월'과 '마감'이 있고, 보낸 시간이 09:26 인 메일
        if ($mail.Subject -like "*2월*마감*" -and $mail.SentOn.ToString("HH:mm") -eq "09:26") {
            $targetMail = $mail
            break
        }
    }
    
    if ($targetMail -eq $null) {
        # 못 찾으면 가장 첫번째(최신) 메일이라도 시도
        $targetMail = $items.Item(1)
    }

    $recipientsAddresses = @()
    foreach ($recip in $targetMail.Recipients) {
        if ($recip.Type -eq 1) { # olTo
            $recipientsAddresses += $recip.Address
        }
    }

    $subject = [char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]54869 + [char]51064
    $body = [char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]50756 + [char]47308 + [char]32 + [char]49324 + [char]54637 + [char]32 + [char]54869 + [char]51064 + [char]51012 + [char]32 + [char]50948 + [char]54620 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]51077 + [char]45768 + [char]45796 + [char]46
    
    $appt = $outlook.CreateItem(1)
    $appt.Subject = $subject
    $appt.Body = $body
    $appt.Location = "https://sskv.webex.com/sskv/j.php?MTID=mb9e45aa44dbb35ccfc41197aabb63e1b"
    $appt.Start = Get-Date -Year 2026 -Month 3 -Day 5 -Hour 10 -Minute 30 -Second 0
    $appt.End = Get-Date -Year 2026 -Month 3 -Day 5 -Hour 12 -Minute 0 -Second 0
    $appt.MeetingStatus = 1

    foreach ($addr in $recipientsAddresses) {
        $newRecip = $appt.Recipients.Add($addr)
        $newRecip.Type = 1
    }
    $appt.Recipients.ResolveAll()
    $appt.Display()
    "SUCCESS"
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\last_hope.log"
    throw $_
}
