
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $sentFolder = $namespace.GetDefaultFolder(5)
    
    # Sort and get the very last item directly from the folder
    $latestMail = $sentFolder.Items | Sort-Object SentOn -Descending | Select-Object -First 1

    if ($latestMail -eq $null) {
        throw "보낸 메일을 찾을 수 없습니다."
    }

    $recipients = @()
    foreach ($recip in $latestMail.Recipients) {
        if ($recip.Type -eq 1) { # olTo
            $recipients += $recip.Address
        }
    }

    # '2월 마감 회의 확인' 유니코드 조립
    $subject = [char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]54869 + [char]51064
    
    $appt = $outlook.CreateItem(1)
    $appt.Subject = $subject
    $appt.MeetingStatus = 1
    $appt.Start = (Get-Date).AddHours(1).ToString("yyyy-MM-dd HH:00:00")
    $appt.Duration = 60
    
    foreach ($addr in $recipients) {
        $appt.Recipients.Add($addr)
    }

    # 화면에 표시
    $appt.Display()
    
    "SUCCESS: Displayed meeting for " + ($recipients -join ", ")
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\last_chance.log"
    throw $_
}
